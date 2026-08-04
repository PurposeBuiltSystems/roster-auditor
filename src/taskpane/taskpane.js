/*
 * Roster Auditor — task pane UI wiring.
 *
 * Flow: load groups -> pick one -> audit (roster + NDRs + group sends +
 * replies + Teams attendance for meetings the user organized + manual
 * attendance) -> review rows -> re-engage / remove / export CSV.
 *
 * Removal is two-click (arm, then confirm) and only offered on M365 groups
 * the user owns; for Graph-read-only groups (classic DLs, mail-enabled
 * security) the deliverable is the removal-recommendation CSV.
 */
/* global Office, GraphData, RosterAuditor, document, window, Blob, URL */
(function () {
  "use strict";

  var MANUAL_KEY = "rosteraud.manual"; // { groupId: { email: [YYYY-MM-DD] } }
  var LOG_KEY = "rosteraud.log";
  var MANUAL_DATES_PER_MEMBER = 8;
  var MANUAL_GROUP_CAP = 6;
  var LOG_CAP = 40;

  var groups = [];
  var ownedIds = {};
  var current = null; // { group, meta }
  var rows = [];
  var removeArmed = false;

  Office.onReady(function () {
    on("loadGroups", "click", loadGroups);
    on("audit", "click", audit);
    on("groupSelect", "change", groupChanged);
    on("draftBtn", "click", draftReengage);
    on("removeBtn", "click", removeSelected);
    on("csvSelBtn", "click", function () { exportCsv(true); });
    on("csvAllBtn", "click", function () { exportCsv(false); });
    byId("manualDate").value = new Date().toISOString().slice(0, 10);
    renderLog();
  });

  function byId(id) { return document.getElementById(id); }

  /**
   * Outlook caches the pane HTML but the ?v= query string makes it fetch
   * JavaScript fresh, so a returning user can run today's JS against
   * yesterday's page. Binding through this helper means a missing element
   * costs one feature instead of throwing and leaving every later button
   * unbound — a whole dead pane.
   */
  function on(id, ev, fn) {
    var el = byId(id);
    if (el) { el.addEventListener(ev, fn); }
    return el;
  }

  function setStatus(kind, text) {
    var el = byId("status");
    if (!text) { el.hidden = true; return; }
    el.hidden = false;
    el.className = "status " + kind;
    el.textContent = text;
  }

  /* ------------------------------------------------------ roaming settings */

  function readJson(key, fallback) {
    try {
      var raw = Office.context.roamingSettings.get(key);
      return raw ? JSON.parse(raw) : fallback;
    } catch (e) { return fallback; }
  }

  function writeJson(key, value) {
    try {
      Office.context.roamingSettings.set(key, JSON.stringify(value));
      Office.context.roamingSettings.saveAsync(function () {});
    } catch (e) { /* session-only */ }
  }

  function readManual(groupId) {
    var all = readJson(MANUAL_KEY, {});
    return all[groupId] || {};
  }

  function saveManualDate(groupId, email, dateIso) {
    var all = readJson(MANUAL_KEY, {});
    // roamingSettings has a 32 KB TOTAL cap for the whole add-in — keep the
    // stored history small: recent dates per member, a handful of groups.
    var keys = Object.keys(all);
    if (!all[groupId] && keys.length >= MANUAL_GROUP_CAP) { delete all[keys[0]]; }
    var g = all[groupId] || (all[groupId] = {});
    var dates = g[email] || (g[email] = []);
    if (dates.indexOf(dateIso) === -1) {
      dates.push(dateIso);
      dates.sort();
      if (dates.length > MANUAL_DATES_PER_MEMBER) { dates.splice(0, dates.length - MANUAL_DATES_PER_MEMBER); }
    }
    writeJson(MANUAL_KEY, all);
  }

  function log(text) {
    var entries = readJson(LOG_KEY, []);
    entries.push(new Date().toISOString().slice(0, 16).replace("T", " ") + " — " + text);
    if (entries.length > LOG_CAP) { entries.splice(0, entries.length - LOG_CAP); }
    writeJson(LOG_KEY, entries);
    renderLog();
  }

  function renderLog() {
    var host = byId("auditLog");
    var entries = readJson(LOG_KEY, []);
    host.innerHTML = "";
    if (!entries.length) { host.textContent = "No actions recorded yet."; return; }
    entries.slice(-12).reverse().forEach(function (e) {
      var d = document.createElement("div");
      d.textContent = e;
      host.appendChild(d);
    });
  }

  /* ----------------------------------------------------------------- groups */

  async function loadGroups() {
    byId("loadGroups").disabled = true;
    setStatus("work", "Loading the groups you belong to…");
    try {
      var token = await GraphData.getToken();
      var res = await GraphData.myGroups(token);
      groups = res.groups;
      ownedIds = res.ownedIds;
      var sel = byId("groupSelect");
      sel.innerHTML = "";
      groups.forEach(function (g, i) {
        var opt = document.createElement("option");
        opt.value = String(i);
        opt.textContent = g.displayName + (g.mail ? " <" + g.mail + ">" : "");
        sel.appendChild(opt);
      });
      if (!groups.length) {
        setStatus("info", "No groups found on your account.");
        return;
      }
      byId("groupRow").hidden = false;
      groupChanged();
      setStatus("info", groups.length + " groups loaded. Pick one and audit.");
    } catch (e) {
      var msg = (e && e.message) || String(e);
      if (/REPLACE_WITH_ENTRA_CLIENT_ID/.test(GraphData._config.clientId)) {
        msg = "Set your Entra client ID in src/graph.js before running. (" + msg + ")";
      }
      setStatus("error", "Could not load groups: " + msg);
    } finally {
      byId("loadGroups").disabled = false;
    }
  }

  function groupChanged() {
    var g = groups[parseInt(byId("groupSelect").value, 10) || 0];
    if (!g) { return; }
    var meta = RosterAuditor.classifyGroup(g, ownedIds);
    current = { group: g, meta: meta };
    byId("groupKind").textContent = meta.label;
    byId("audit").disabled = false;
    byId("results").hidden = true;
    rows = [];
  }

  /* ------------------------------------------------------------------ audit */

  async function audit() {
    if (!current) { return; }
    var months = Math.max(1, Math.min(12, parseInt(byId("monthsBack").value, 10) || 6));
    var daysBack = months * 30;
    var largeSendMin = Math.max(2, Math.min(20, parseInt(byId("largeSendMin").value, 10) || 3));
    byId("audit").disabled = true;
    byId("results").hidden = true;
    removeArmed = false;
    try {
      var token = await GraphData.getToken();

      setStatus("work", "Reading the roster…");
      var roster = await GraphData.groupMembers(token, current.group.id);

      setStatus("work", "Scanning for bounced sends (NDRs)…");
      var ndrs = await GraphData.ndrMessages(token, daysBack);

      setStatus("work", "Scanning your sent mail for group sends…");
      var sent = await GraphData.sentMessages(token, daysBack);

      setStatus("work", "Scanning replies…");
      var inbound = await GraphData.inboundMessages(token, daysBack);

      setStatus("work", "Finding Teams meetings you organized…");
      var events = await GraphData.myOnlineMeetings(token, daysBack);
      var meetings = [];
      for (var i = 0; i < events.length; i++) {
        setStatus("work", "Attendance " + (i + 1) + "/" + events.length + ": " + (events[i].subject || "(no subject)"));
        var got = await GraphData.attendanceForEvent(token, events[i]);
        meetings = meetings.concat(got);
      }

      rows = RosterAuditor.buildParticipation({
        roster: roster,
        ndrs: ndrs,
        sent: sent,
        inbound: inbound,
        meetings: meetings,
        manual: readManual(current.group.id),
        opts: { groupAddress: current.group.mail || "", largeSendMin: largeSendMin },
      });

      render();
      byId("results").hidden = false;
      var nb = rows.filter(function (r) { return r.status === "bounced"; }).length;
      var nu = rows.filter(function (r) { return r.status === "unknown"; }).length;
      var sends = rows.length ? rows[0].groupSendCount : 0;
      setStatus("info", rows.length + " members · " + nb + " bouncing · " + nu +
        " engagement unknown · evidence: " + sends + " group sends, " +
        meetings.length + " attendance reports, " + ndrs.length + " NDRs (" + months + " mo).");
    } catch (e) {
      setStatus("error", "Audit failed: " + ((e && e.message) || e));
    } finally {
      byId("audit").disabled = false;
    }
  }

  /* ----------------------------------------------------------------- render */

  function statusChip(r) {
    if (r.status === "bounced") {
      return { cls: "chip-bounce", text: "⚠️ bouncing since " + r.bounce.date.slice(0, 10) };
    }
    if (r.status === "active") {
      var label = { manual: "in person", teams: "Teams meeting", reply: "replied" }[r.lastSeen.signal];
      return { cls: "chip-active", text: "✓ " + label + " " + r.lastSeen.date.slice(0, 10) };
    }
    return { cls: "chip-unknown", text: "? engagement unknown" };
  }

  function selected() {
    return rows.filter(function (r) { return r.checked && !r.gone; });
  }

  function updateActions() {
    var n = selected().length;
    var w = current && current.meta.writable;
    byId("draftBtn").textContent = "✉️ Re-engage (" + n + ")";
    byId("draftBtn").disabled = !n;
    var rb = byId("removeBtn");
    rb.textContent = removeArmed ? "Confirm remove (" + n + ")" : "🗑 Remove (" + n + ")";
    rb.className = removeArmed ? "danger-armed" : "";
    rb.disabled = !n || !w;
    rb.title = w ? "Remove the selected members from this Microsoft 365 group"
      : "This group type is read-only via Graph — use the CSV export for your Exchange admin.";
    byId("csvSelBtn").textContent = "CSV (" + n + ")";
    byId("csvSelBtn").disabled = !n;
  }

  function render() {
    var host = byId("items");
    host.innerHTML = "";
    rows.forEach(function (r) {
      var div = document.createElement("div");
      div.className = "member" + (r.gone ? " gone" : "");

      var cb = document.createElement("input");
      cb.type = "checkbox";
      cb.checked = !!r.checked;
      cb.disabled = !!r.gone;
      cb.addEventListener("change", function () {
        r.checked = cb.checked;
        removeArmed = false;
        updateActions();
      });

      var who = document.createElement("div");
      who.className = "who";
      var name = document.createElement("div");
      name.className = "name";
      name.textContent = r.name;
      var email = document.createElement("div");
      email.className = "email";
      email.textContent = r.email || "(no email address)";
      var chips = document.createElement("div");
      chips.className = "chips";
      var chip = statusChip(r);
      var c = document.createElement("span");
      c.className = chip.cls;
      c.textContent = chip.text;
      chips.appendChild(c);
      if (r.gone) {
        var g = document.createElement("span");
        g.textContent = "removed";
        chips.appendChild(g);
      }
      who.appendChild(name);
      who.appendChild(email);
      who.appendChild(chips);

      var side = document.createElement("div");
      side.className = "side";
      var present = document.createElement("button");
      present.className = "present";
      present.textContent = "Present ✓";
      present.title = "Record in-person attendance for the date picked above";
      present.disabled = !!r.gone || !r.email;
      present.addEventListener("click", function () {
        var date = byId("manualDate").value || new Date().toISOString().slice(0, 10);
        saveManualDate(current.group.id, r.email, date);
        if (r.status !== "bounced") {
          if (!r.lastSeen || date > r.lastSeen.date.slice(0, 10)) {
            r.lastSeen = { date: date, signal: "manual" };
            r.status = "active";
          }
        }
        log("Marked " + r.email + " present on " + date + " (" + current.group.displayName + ")");
        render();
      });
      side.appendChild(present);

      div.appendChild(cb);
      div.appendChild(who);
      div.appendChild(side);
      host.appendChild(div);
    });
    updateActions();
  }

  /* ---------------------------------------------------------------- actions */

  async function draftReengage() {
    var sel = selected();
    if (!sel.length) { return; }
    byId("draftBtn").disabled = true;
    setStatus("work", "Creating the re-engagement draft…");
    try {
      var token = await GraphData.getToken();
      var gname = current.group.displayName;
      var subject = "Are you still an active member of “" + gname + "”?";
      var body =
        "<p>Hi,</p>" +
        "<p>I'm tidying up the membership roster for <b>" + gname + "</b> and I " +
        "haven't seen recent participation from you — that may simply mean the " +
        "list has your details wrong, or that you follow along without replying.</p>" +
        "<p>Could you reply to let me know whether you'd like to stay on the " +
        "roster? If I don't hear back, I may remove you to keep the list current — " +
        "you'd always be welcome back.</p>" +
        "<p>Thanks!</p>";
      var draft = await GraphData.createDraft(token, subject, body,
        sel.map(function (r) { return r.email; }).filter(Boolean));
      log("Re-engagement draft for " + sel.length + " member(s) of " + gname);
      setStatus("info", "Draft created for " + sel.length + " member(s) — check your Drafts folder if it didn't open.");
      try {
        var ewsId = Office.context.mailbox.convertToEwsId(draft.id, Office.MailboxEnums.RestVersion.v2_0);
        Office.context.mailbox.displayMessageForm(ewsId);
      } catch (openErr) { /* the status line covers it */ }
    } catch (e) {
      setStatus("error", "Could not create the draft: " + ((e && e.message) || e));
    } finally {
      byId("draftBtn").disabled = false;
      updateActions();
    }
  }

  async function removeSelected() {
    var sel = selected();
    if (!sel.length || !current.meta.writable) { return; }
    if (!removeArmed) { // first click arms; second click executes
      removeArmed = true;
      updateActions();
      setStatus("info", "Click “Confirm remove” to remove " + sel.length +
        " member(s) from " + current.group.displayName + ". This is immediate.");
      return;
    }
    removeArmed = false;
    byId("removeBtn").disabled = true;
    var ok = 0, failed = [];
    try {
      var token = await GraphData.getToken();
      for (var i = 0; i < sel.length; i++) {
        setStatus("work", "Removing " + (i + 1) + "/" + sel.length + ": " + sel[i].email);
        try {
          await GraphData.removeMember(token, current.group.id, sel[i].id);
          sel[i].gone = true;
          sel[i].checked = false;
          ok++;
          log("Removed " + sel[i].email + " from " + current.group.displayName);
        } catch (e) {
          failed.push(sel[i].email);
        }
      }
      render();
      if (failed.length) {
        setStatus("error", "Removed " + ok + "; failed for: " + failed.join(", ") +
          ". (Owner rights on the group are required.)");
      } else {
        setStatus("info", "Removed " + ok + " member(s) from " + current.group.displayName + ".");
      }
    } catch (e) {
      setStatus("error", "Removal failed: " + ((e && e.message) || e));
    } finally {
      updateActions();
    }
  }

  function exportCsv(selectedOnly) {
    var list = selectedOnly ? selected() : rows;
    if (!list.length) { return; }
    var csv = RosterAuditor.buildRemovalCsv(list, current.group.displayName);
    var fname = "roster-audit-" +
      current.group.displayName.replace(/[^A-Za-z0-9-]+/g, "_").slice(0, 40) +
      "-" + new Date().toISOString().slice(0, 10) + ".csv";
    try {
      var blob = new Blob([csv], { type: "text/csv" });
      var a = document.createElement("a");
      a.href = URL.createObjectURL(blob);
      a.download = fname;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      setStatus("info", "CSV downloaded (" + list.length + " rows) — includes a ready-made Exchange PowerShell line per member.");
    } catch (e) {
      // webview without download support: show the CSV to copy out
      var ta = document.createElement("textarea");
      ta.className = "csv-out";
      ta.value = csv;
      byId("items").insertBefore(ta, byId("items").firstChild);
      ta.select();
      setStatus("info", "Copy the CSV from the box above (this client blocks downloads).");
    }
    log("Exported CSV (" + list.length + " rows) for " + current.group.displayName);
  }
})();
