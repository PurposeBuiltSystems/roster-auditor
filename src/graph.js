/*
 * Roster Auditor — Microsoft Graph data layer.
 *
 * AUTH: Nested App Authentication (NAA) via MSAL — no backend; identical
 * pattern to the other PurposeBuilt add-ins. Everything is DELEGATED and
 * limited to what the signed-in user can already see: their own mail (NDRs,
 * group sends, replies), attendance reports for meetings THEY organized,
 * and rosters of groups they belong to or own. Removal writes only work on
 * Microsoft 365 groups the user owns — classic DLs are read-only via Graph
 * by platform design (Exchange-mastered), which is why the pane exports a
 * CSV for those instead.
 *
 * Exposes a global `GraphData` object.
 */
/* global msal */
(function (root) {
  "use strict";

  var CLIENT_ID = "a29d3a83-ea45-47ca-b628-3a8ec45cd87d"; // "Roster Auditor" Entra app (purposebuilt.systems tenant)
  var GRAPH = "https://graph.microsoft.com/v1.0";
  var SCOPES = [
    "Group.Read.All",              // list groups + members
    "Group.ReadWrite.All",         // remove members (M365 groups the user owns)
    "Mail.ReadWrite",              // NDR scan, group-send scan, re-engagement draft
    "Calendars.Read",              // enumerate the user's meetings
    "OnlineMeetings.Read",         // resolve calendar event -> onlineMeeting id
    "OnlineMeetingArtifact.Read.All", // attendance reports (organizer-only)
  ];

  var pcaPromise = null;

  function getPca() {
    if (!pcaPromise) {
      pcaPromise = msal.createNestablePublicClientApplication({
        auth: {
          clientId: CLIENT_ID,
          authority: "https://login.microsoftonline.com/common",
        },
      });
    }
    return pcaPromise;
  }

  /**
   * Signed-in account, or null. Lets the pane say who is signed in and
   * decide whether Sign out has anything to act on.
   */
  async function currentAccount() {
    try {
      var pca = await getPca();
      var accts = (pca.getAllAccounts && pca.getAllAccounts()) || [];
      return accts.length ? (accts[0].username || accts[0].name || "signed in") : null;
    } catch (e) { return null; }
  }

  /**
   * Sign out of the ADD-IN. Certification policy 1100.5.7.1 requires a way
   * out wherever there is a way in. Under nested app authentication Outlook
   * owns the session and no add-in can end it, so this clears every token
   * and account this add-in has cached - and the pane says exactly that
   * rather than implying more than it does.
   */
  async function signOut() {
    var pca = null;
    try { pca = await getPca(); } catch (e) { /* nothing cached to clear */ }
    if (pca) {
      var accts = (pca.getAllAccounts && pca.getAllAccounts()) || [];
      for (var i = 0; i < accts.length; i++) {
        try {
          if (pca.clearCache) { await pca.clearCache({ account: accts[i] }); }
          else if (pca.logoutPopup) { await pca.logoutPopup({ account: accts[i] }); }
        } catch (e) { /* keep clearing the rest */ }
      }
    }
    pcaPromise = null;   // force a fresh broker handshake next time
    return true;
  }


  /**
   * Sign-in must never hang the pane. An un-timed await on the popup flow
   * leaves a button disabled with nothing visible happening — which reads
   * to the user as "the button does nothing" and gives them nothing to act
   * on. Fail loudly instead, naming the two things that actually fix it.
   */
  function withTimeout(promise, ms, message) {
    var timer;
    return Promise.race([
      promise.then(function (v) { clearTimeout(timer); return v; },
                   function (e) { clearTimeout(timer); throw e; }),
      new Promise(function (_, reject) {
        timer = setTimeout(function () { reject(new Error(message)); }, ms);
      }),
    ]);
  }

  async function getToken() {
    var pca = await withTimeout(getPca(), 20000,
      "Sign-in didn't start. Fully quit Outlook (Cmd+Q) and reopen, then try again.");
    try {
      return (await withTimeout(pca.acquireTokenSilent({ scopes: SCOPES }), 20000, "silent timeout")).accessToken;
    } catch (e) {
      var interactive = await withTimeout(
        pca.acquireTokenPopup({ scopes: SCOPES }), 120000,
        "Sign-in didn't finish. A Microsoft sign-in window may have opened behind Outlook — " +
        "check for it (or Mission Control), finish signing in, and click again. If no window " +
        "appeared at all, fully quit Outlook (Cmd+Q), reopen, and retry.");
      return interactive.accessToken;
    }
  }

  async function graph(token, method, path, body) {
    var res = await fetch(GRAPH + path, {
      method: method,
      headers: {
        Authorization: "Bearer " + token,
        "Content-Type": "application/json",
      },
      body: body ? JSON.stringify(body) : undefined,
    });
    if (!res.ok) {
      var text = await res.text();
      throw new Error("Graph " + method + " " + path + " -> " + res.status + " " + text);
    }
    return res.status === 204 ? null : res.json();
  }

  /** Page through a Graph collection following @odata.nextLink. */
  async function graphAll(token, path, cap) {
    var items = [];
    var url = GRAPH + path;
    var guard = 0;
    while (url && guard++ < (cap || 30)) {
      var res = await fetch(url, { headers: { Authorization: "Bearer " + token } });
      if (!res.ok) { throw new Error("Graph GET " + url + " -> " + res.status); }
      var page = await res.json();
      items = items.concat(page.value || []);
      url = page["@odata.nextLink"] || null;
    }
    return items;
  }

  /* ---------------------------------------------------------------- groups */

  var GROUP_SELECT = "$select=id,displayName,mail,mailEnabled,securityEnabled,groupTypes";

  /** Groups the user belongs to, plus a set of the ids they OWN. DLs and
   *  mail-enabled security groups do appear in memberOf (read-only). */
  async function myGroups(token) {
    var member = await graphAll(token, "/me/memberOf/microsoft.graph.group?" + GROUP_SELECT + "&$top=999");
    var owned = [];
    try {
      owned = await graphAll(token, "/me/ownedObjects/microsoft.graph.group?" + GROUP_SELECT + "&$top=999");
    } catch (e) { /* ownedObjects can 403 in some tenants; owned stays empty */ }
    var ownedIds = {};
    owned.forEach(function (g) { ownedIds[g.id] = true; });
    var byId = {};
    member.concat(owned).forEach(function (g) { byId[g.id] = g; });
    var groups = Object.keys(byId).map(function (id) { return byId[id]; });
    groups.sort(function (a, b) { return (a.displayName || "").localeCompare(b.displayName || ""); });
    return { groups: groups, ownedIds: ownedIds };
  }

  /** Roster. Generic /members (not the user cast) so DL contacts count too. */
  async function groupMembers(token, groupId) {
    return graphAll(token,
      "/groups/" + groupId + "/members?$select=id,displayName,mail,userPrincipalName&$top=999", 40);
  }

  /** Remove a member — M365 groups the user owns ONLY (Graph rejects DLs). */
  async function removeMember(token, groupId, memberId) {
    return graph(token, "DELETE", "/groups/" + groupId + "/members/" + memberId + "/$ref");
  }

  /* ------------------------------------------------------------------ mail */

  var sinceIso = function (daysBack) {
    return new Date(Date.now() - daysBack * 864e5).toISOString();
  };

  /** Sent messages with recipients, for group-send detection. */
  async function sentMessages(token, daysBack) {
    return graphAll(token,
      "/me/mailFolders/sentitems/messages" +
      "?$select=id,conversationId,subject,sentDateTime,toRecipients,ccRecipients" +
      "&$filter=sentDateTime ge " + sinceIso(daysBack) + "&$top=100");
  }

  /** Everything received in the window (reply signal). */
  async function inboundMessages(token, daysBack) {
    return graphAll(token,
      "/me/messages" +
      "?$select=id,conversationId,receivedDateTime,from" +
      "&$filter=receivedDateTime ge " + sinceIso(daysBack) + "&$top=100");
  }

  /**
   * NDRs: filter on the MAPI message class (String 0x001a = PR_MESSAGE_CLASS)
   * = REPORT.IPM.Note.NDR. Falls back to a subject-prefix match if a tenant
   * rejects the extended-property filter.
   */
  async function ndrMessages(token, daysBack) {
    var select = "$select=id,subject,receivedDateTime,body";
    try {
      return await graphAll(token,
        "/me/messages?" + select +
        "&$filter=receivedDateTime ge " + sinceIso(daysBack) +
        " and singleValueExtendedProperties/Any(ep: ep/id eq 'String 0x001a'" +
        " and ep/value eq 'REPORT.IPM.Note.NDR')&$top=50", 10);
    } catch (e) {
      return graphAll(token,
        "/me/messages?" + select +
        "&$filter=receivedDateTime ge " + sinceIso(daysBack) +
        " and startswith(subject,'Undeliverable')&$top=50", 10);
    }
  }

  /** Re-engagement draft to the selected members; returns {id, webLink}. */
  async function createDraft(token, subject, bodyHtml, toEmails) {
    var draft = await graph(token, "POST", "/me/messages", {
      subject: subject,
      body: { contentType: "HTML", content: bodyHtml },
      toRecipients: toEmails.map(function (e) { return { emailAddress: { address: e } }; }),
    });
    return { id: draft.id, webLink: draft.webLink || "" };
  }

  /* ------------------------------------------------------------- meetings */

  /** Teams meetings the user ORGANIZED in the window (join URL present). */
  async function myOnlineMeetings(token, daysBack) {
    var start = sinceIso(daysBack);
    var end = new Date().toISOString();
    var events = await graphAll(token,
      "/me/calendarView?startDateTime=" + start + "&endDateTime=" + end +
      "&$select=subject,start,isOrganizer,onlineMeeting&$top=100", 20);
    return events.filter(function (ev) {
      return ev.isOrganizer && ev.onlineMeeting && ev.onlineMeeting.joinUrl;
    });
  }

  /**
   * Attendance for one organized meeting, normalized to
   * { date, subject, attendees: [emails] } entries (one per session report).
   * Returns [] quietly when reports are off/absent — the auditor treats
   * missing data as "no signal", never as absence.
   */
  async function attendanceForEvent(token, ev) {
    var joinUrl = ev.onlineMeeting.joinUrl;
    var found;
    try {
      found = await graph(token, "GET",
        "/me/onlineMeetings?$filter=JoinWebUrl eq '" + encodeURIComponent(joinUrl).replace(/'/g, "%27") + "'");
    } catch (e) { return []; }
    var meeting = found && found.value && found.value[0];
    if (!meeting) { return []; }
    var out = [];
    try {
      var reports = await graph(token, "GET",
        "/me/onlineMeetings/" + meeting.id + "/attendanceReports");
      var list = (reports && reports.value) || [];
      for (var i = 0; i < Math.min(list.length, 10); i++) {
        var full = await graph(token, "GET",
          "/me/onlineMeetings/" + meeting.id + "/attendanceReports/" + list[i].id +
          "?$expand=attendanceRecords");
        var recs = (full && full.attendanceRecords) || [];
        out.push({
          date: full.meetingStartDateTime || (ev.start && ev.start.dateTime) || "",
          subject: ev.subject || "",
          attendees: recs.map(function (r) { return (r.emailAddress || "").toLowerCase(); })
            .filter(Boolean),
        });
      }
    } catch (e) { /* attendance disabled or not retained for this meeting */ }
    return out;
  }

  root.GraphData = {
    signOut: signOut,
    currentAccount: currentAccount,
    getToken: getToken,
    myGroups: myGroups,
    groupMembers: groupMembers,
    removeMember: removeMember,
    sentMessages: sentMessages,
    inboundMessages: inboundMessages,
    ndrMessages: ndrMessages,
    createDraft: createDraft,
    myOnlineMeetings: myOnlineMeetings,
    attendanceForEvent: attendanceForEvent,
    _config: { clientId: CLIENT_ID },
  };
})(typeof self !== "undefined" ? self : this);
