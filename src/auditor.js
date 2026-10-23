/*
 * Roster Auditor — pure participation logic. No Office.js, no Graph, no DOM;
 * unit-testable under Node (see test/auditor.test.js).
 *
 * Inputs are plain objects shaped like Graph responses. The auditor
 * cross-references a group roster against four signals, strongest first:
 *
 *   1. bounce  — a delivery-failure diagnostic identifying the member
 *   2. manual  — clerk-recorded in-person attendance
 *   3. teams   — attendance records from meetings the USER organized
 *   4. reply   — an inbound reply on a conversation the user sent to the group
 *                (weak: silence is not absence — display only, never auto-flag)
 *
 * Output rows carry a lastSeen {date, signal} and a status:
 *   'bounced' | 'active' | 'unknown'  — "unknown", never "inactive": all we
 * can honestly say is that none of OUR signals saw them.
 */
(function (root) {
  "use strict";

  var SIGNAL_RANK = { manual: 3, teams: 2, reply: 1 };

  function norm(addr) {
    return String(addr || "").trim().toLowerCase();
  }

  /* ---------------------------------------------------------------- groups */

  /**
   * Classify a Graph group object into what this tool can do with it.
   * CONFIRMED Graph reality: distribution lists and mail-enabled security
   * groups are Exchange-mastered and READ-ONLY via Graph — removal for those
   * goes out as a CSV for the Exchange admin, never a button.
   */
  function classifyGroup(g, ownedIds) {
    var types = g.groupTypes || [];
    var unified = types.indexOf("Unified") !== -1;
    var dynamic = types.indexOf("DynamicMembership") !== -1;
    var owned = !!(ownedIds && ownedIds[g.id]);
    var kind, label, writable = false;

    if (dynamic) {
      kind = "dynamic";
      label = "Dynamic membership (rule-based — edit the rule, not members)";
    } else if (unified) {
      kind = "m365";
      writable = owned;
      label = owned ? "Microsoft 365 group (you own it — confirmed removal)"
                    : "Microsoft 365 group (not owner — export CSV)";
    } else if (g.mailEnabled && g.securityEnabled) {
      kind = "mailSecurity";
      label = "Mail-enabled security group (read-only via Graph — export CSV)";
    } else if (g.mailEnabled) {
      kind = "dl";
      label = "Distribution list (read-only via Graph — export CSV)";
    } else {
      kind = "security";
      label = "Security group (export CSV)";
    }
    return { kind: kind, writable: writable, owned: owned, label: label };
  }

  /* ---------------------------------------------------------------- NDRs */

  var EMAIL_RE = /[A-Za-z0-9._%+'-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}/g;
  var NDR_NOISE = /^(postmaster|mailer-daemon|microsoftexchange|no-?reply)/;

  function stripHtml(html) {
    return String(html || "")
      .replace(/<style[\s\S]*?<\/style>/gi, " ")
      .replace(/<br\s*\/?\s*>|<\/(?:p|div|tr|li)>/gi, "\n")
      .replace(/<[^>]+>/g, " ")
      .replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&amp;/g, "&")
      .replace(/&#39;/g, "'").replace(/&quot;/g, '"').replace(/&nbsp;/g, " ");
  }

  /**
   * Read failed-recipient diagnostics, not arbitrary roster addresses in an
   * NDR. Quoted original recipients can include people whose delivery worked.
   * Unrecognized layouts produce no bounce evidence rather than a guess.
   * ndr: { subject, receivedDateTime, body: { content } } (bodyPreview ok too)
   */
  function ndrFailedRecipients(ndr, rosterEmails) {
    var rosterSet = {};
    (rosterEmails || []).forEach(function (e) { rosterSet[norm(e)] = true; });
    var text = stripHtml((ndr.body && ndr.body.content) || ndr.bodyPreview || "")
      .replace(/\r\n?/g, "\n");
    // A forwarded original is evidence of what was sent, not what failed.
    text = text.split(/\bOriginal (?:message(?: headers)?|headers)\s*:|Content-Type:\s*message\/rfc822|-{3,}\s*(?:Original|Forwarded) message\s*-{3,}/i)[0];
    var seen = {};
    var hits = [];
    function add(raw) {
      var e = norm(raw);
      if (seen[e] || NDR_NOISE.test(e) || !rosterSet[e]) { return; }
      seen[e] = true;
      hits.push(e);
    }
    // RFC 3464: each Final-Recipient starts a new recipient record. Only
    // failed records count; delivered/delayed recipients must stay unflagged.
    if (/^Final-Recipient:/im.test(text)) {
      text.split(/(?=^Final-Recipient:)/im).forEach(function (block) {
        var recipient = block.match(/^Final-Recipient:\s*[^;\n]+;\s*([^\n]+)/im);
        if (!recipient || !/^Action:\s*failed\s*$/im.test(block)) { return; }
        (recipient[1].match(EMAIL_RE) || []).forEach(add);
      });
      return hits;
    }
    var failureList = false;
    text.split("\n").forEach(function (line) {
      line = line.trim();
      if (/^(?:Diagnostic information|Generating server|For (?:email )?admins|From:|To:|Cc:|Subject:)/i.test(line)) {
        failureList = false;
      }
      // Outlook's human-readable NDR starts with a failure heading, then
      // recipient-only lines. Do not collect addresses from explanatory prose.
      if (/^Delivery (?:has )?failed to (?:these|the following) recipients(?: or groups)?\s*:/i.test(line)) {
        failureList = true;
        line = line.replace(/^Delivery (?:has )?failed to (?:these|the following) recipients(?: or groups)?\s*:/i, "").trim();
      }
      if (failureList && line) {
        var addresses = line.match(EMAIL_RE) || [];
        var remainder = line.replace(EMAIL_RE, "");
        // Accept bare recipient lists or a single display-name address.
        if (addresses.length && (/^[\s<>(),;]*$/.test(remainder) ||
            (addresses.length === 1 && /^[^<>:;]+<[^<>]+>$/.test(line)))) {
          addresses.forEach(add); return;
        }
      }
      if (line) { failureList = false; }
      // Short diagnostics put the recipient immediately before/after the
      // failure phrase. Match that recipient only, never the whole sentence.
      var address = "([A-Za-z0-9._%+'-]+@[A-Za-z0-9.-]+\\.[A-Za-z]{2,})";
      var before = new RegExp("(?:delivery (?:has )?failed(?: to)?|failed for)\\s*:?\\s*<?" + address, "gi");
      var after = new RegExp("(?:Recipient\\s+)?<?" + address + ">?\\s+(?:could not be reached|failed\\b|does not exist\\b|rejected\\b)", "gi");
      var match;
      while ((match = before.exec(line))) { add(match[1]); }
      while ((match = after.exec(line))) { add(match[1]); }
    });
    return hits;
  }

  /* ------------------------------------------------------------ group sends */

  function recipients(msg) {
    return (msg.toRecipients || []).concat(msg.ccRecipients || [])
      .map(function (r) { return norm(r.emailAddress && r.emailAddress.address); })
      .filter(Boolean);
  }

  /**
   * Which of the user's sent messages went "to the group"? Either addressed
   * to the group's own address, or a large send reaching >= largeSendMin
   * roster members directly.
   */
  function groupSends(sent, opts) {
    var groupAddr = norm(opts.groupAddress);
    var rosterSet = {};
    (opts.rosterEmails || []).forEach(function (e) { rosterSet[norm(e)] = true; });
    var min = opts.largeSendMin || 3;

    return (sent || []).filter(function (m) {
      var rcpts = recipients(m);
      if (groupAddr && rcpts.indexOf(groupAddr) !== -1) { return true; }
      var n = 0;
      rcpts.forEach(function (r) { if (rosterSet[r]) { n++; } });
      return n >= min;
    });
  }

  /* -------------------------------------------------------- participation */

  /**
   * The core cross-reference. Returns one row per roster member, sorted for
   * review: bounced first, then never-seen, then active stalest-first.
   *
   * roster:   [{ id, displayName, mail, userPrincipalName }]
   * ndrs:     [{ subject, receivedDateTime, body|bodyPreview }]
   * sent:     user's sent messages (Graph shape, with recipients)
   * inbound:  [{ conversationId, receivedDateTime, from }]
   * meetings: [{ date, subject, attendees: [emails] }]   (Teams attendance, normalized)
   * manual:   { email: [isoDateStrings] }                 (clerk-recorded)
   * opts:     { groupAddress, largeSendMin, now }
   */
  function buildParticipation(input) {
    var roster = input.roster || [];
    var opts = input.opts || {};
    var rosterEmails = roster.map(function (m) {
      return norm(m.mail || m.userPrincipalName);
    }).filter(Boolean);

    // 1. bounces
    var bounced = {}; // email -> { date, subject }
    (input.ndrs || []).forEach(function (ndr) {
      ndrFailedRecipients(ndr, rosterEmails).forEach(function (e) {
        var prev = bounced[e];
        if (!prev || ndr.receivedDateTime > prev.date) {
          bounced[e] = { date: ndr.receivedDateTime, subject: ndr.subject || "" };
        }
      });
    });

    // 2+3. attendance (manual entries and Teams records share one shape)
    var seenAt = {}; // email -> { date, signal }
    function record(email, dateIso, signal) {
      var e = norm(email);
      if (!e || !dateIso) { return; }
      var prev = seenAt[e];
      if (!prev || dateIso > prev.date ||
          (dateIso === prev.date && SIGNAL_RANK[signal] > SIGNAL_RANK[prev.signal])) {
        seenAt[e] = { date: dateIso, signal: signal };
      }
    }
    (input.meetings || []).forEach(function (mtg) {
      (mtg.attendees || []).forEach(function (a) { record(a, mtg.date, "teams"); });
    });
    var manual = input.manual || {};
    Object.keys(manual).forEach(function (email) {
      (manual[email] || []).forEach(function (d) { record(email, d, "manual"); });
    });

    // 4. replies to the user's group sends
    var sends = groupSends(input.sent, {
      groupAddress: opts.groupAddress,
      rosterEmails: rosterEmails,
      largeSendMin: opts.largeSendMin,
    });
    var sendTimeByConv = {};
    sends.forEach(function (s) {
      var t = sendTimeByConv[s.conversationId];
      if (!t || s.sentDateTime < t) { sendTimeByConv[s.conversationId] = s.sentDateTime; }
    });
    (input.inbound || []).forEach(function (msg) {
      var t = sendTimeByConv[msg.conversationId];
      if (!t || !(msg.receivedDateTime > t)) { return; }
      var from = norm(msg.from && msg.from.emailAddress && msg.from.emailAddress.address);
      if (from) { record(from, msg.receivedDateTime, "reply"); }
    });

    // assemble rows
    var rows = roster.map(function (m) {
      var email = norm(m.mail || m.userPrincipalName);
      var b = bounced[email] || null;
      var last = seenAt[email] || null;
      // Later participation is stronger evidence than an older failed send.
      var currentBounce = b && !(last && last.date > b.date);
      return {
        id: m.id,
        name: m.displayName || email,
        email: email,
        bounce: b,
        lastSeen: last,
        status: currentBounce ? "bounced" : (last ? "active" : "unknown"),
        groupSendCount: sends.length,
      };
    });

    var statusOrder = { bounced: 0, unknown: 1, active: 2 };
    rows.sort(function (a, b) {
      if (statusOrder[a.status] !== statusOrder[b.status]) {
        return statusOrder[a.status] - statusOrder[b.status];
      }
      if (a.status === "active") { // stalest first
        return a.lastSeen.date < b.lastSeen.date ? -1 :
               a.lastSeen.date > b.lastSeen.date ? 1 : 0;
      }
      return a.name.localeCompare(b.name);
    });
    return rows;
  }

  /* ------------------------------------------------------------------ CSV */

  function csvEscape(v) {
    var s = String(v == null ? "" : v);
    // Keep untrusted text from becoming a formula when opened in Excel.
    if (typeof v !== "number" && /^[\s\uFEFF]*[=+@-]/.test(s)) { s = "'" + s; }
    return /[",\n\r]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s;
  }

  /**
   * Removal-recommendation CSV — the deliverable for read-only groups
   * (classic DLs / mail-enabled security groups). Includes a ready-made
   * Exchange Online PowerShell line per row so the admin's job is paste-level.
   */
  function buildRemovalCsv(rows, groupName) {
    var header = ["email", "name", "status", "last_seen", "signal",
      "bounce_date", "bounce_subject", "exchange_powershell"];
    var lines = [header.join(",")];
    var psq = function (s) { return "'" + String(s).replace(/'/g, "''") + "'"; }; // literal PowerShell arguments: no variable or command expansion
    (rows || []).forEach(function (r) {
      var ps = "Remove-DistributionGroupMember -Identity " + psq(groupName) +
        " -Member " + psq(r.email) + " -Confirm:$false";
      lines.push([
        r.email,
        r.name,
        r.status,
        r.lastSeen ? r.lastSeen.date.slice(0, 10) : "never seen",
        r.lastSeen ? r.lastSeen.signal : "",
        r.bounce ? r.bounce.date.slice(0, 10) : "",
        r.bounce ? r.bounce.subject : "",
        ps,
      ].map(csvEscape).join(","));
    });
    return lines.join("\r\n") + "\r\n";
  }

  root.RosterAuditor = {
    classifyGroup: classifyGroup,
    ndrFailedRecipients: ndrFailedRecipients,
    groupSends: groupSends,
    buildParticipation: buildParticipation,
    buildRemovalCsv: buildRemovalCsv,
    _internals: { stripHtml: stripHtml, norm: norm, csvEscape: csvEscape },
  };
  if (typeof module !== "undefined" && module.exports) {
    module.exports = root.RosterAuditor;
  }
})(typeof self !== "undefined" ? self : this);
