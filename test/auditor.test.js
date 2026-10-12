/* Offline unit tests for the roster participation auditor. Run: npm test */
"use strict";
var A = require("../src/auditor.js");

var failures = 0;
function check(label, actual, expected) {
  if (JSON.stringify(actual) !== JSON.stringify(expected)) {
    failures++;
    console.error("FAIL  " + label + "\n  expected: " + JSON.stringify(expected) + "\n  actual:   " + JSON.stringify(actual));
  }
}

/* ----------------------------------------------------- group classification */

check("M365 owned -> writable",
  A.classifyGroup({ id: "g1", groupTypes: ["Unified"], mailEnabled: true, securityEnabled: false }, { g1: true }).writable,
  true);
check("M365 not owned -> not writable",
  A.classifyGroup({ id: "g1", groupTypes: ["Unified"], mailEnabled: true, securityEnabled: false }, {}).writable,
  false);
check("classic DL -> kind dl, read-only",
  (function () {
    var m = A.classifyGroup({ id: "g2", groupTypes: [], mailEnabled: true, securityEnabled: false }, { g2: true });
    return [m.kind, m.writable];
  })(),
  ["dl", false]);
check("mail-enabled security -> read-only",
  A.classifyGroup({ id: "g3", groupTypes: [], mailEnabled: true, securityEnabled: true }, { g3: true }).kind,
  "mailSecurity");
check("dynamic wins over unified",
  A.classifyGroup({ id: "g4", groupTypes: ["Unified", "DynamicMembership"], mailEnabled: true }, { g4: true }).kind,
  "dynamic");

/* ---------------------------------------------------------------- NDR parse */

var roster = ["ann@example.com", "bob@example.com", "carol@agency.gov"];

check("NDR html body -> roster hit",
  A.ndrFailedRecipients({
    subject: "Undeliverable: Committee agenda",
    body: { content: "<p>Delivery has failed to these recipients:</p><p><b>ann@example.com</b></p><p>postmaster@relay.example.net</p>" },
  }, roster),
  ["ann@example.com"]);

check("NDR ignores non-roster + noise addresses",
  A.ndrFailedRecipients({
    subject: "Undeliverable: hi",
    body: { content: "failed for stranger@other.org, MAILER-DAEMON@mx.com, no-reply@x.io" },
  }, roster),
  []);

check("NDR entity-encoded angle brackets still parse",
  A.ndrFailedRecipients({
    subject: "Undeliverable: minutes",
    body: { content: "Recipient &lt;bob@example.com&gt; could not be reached." },
  }, roster),
  ["bob@example.com"]);

check("NDR dedupes repeated mention",
  A.ndrFailedRecipients({
    subject: "Undeliverable: x",
    body: { content: "carol@agency.gov failed. carol@agency.gov does not exist." },
  }, roster),
  ["carol@agency.gov"]);

check("NDR case-insensitive match",
  A.ndrFailedRecipients({ subject: "Undeliverable: y", bodyPreview: "Ann@Example.COM rejected" }, roster),
  ["ann@example.com"]);

/* -------------------------------------------------------------- group sends */

function to(addr) { return { emailAddress: { address: addr } }; }

var sent = [
  { id: "s1", conversationId: "c1", sentDateTime: "2026-07-01T10:00:00Z",
    toRecipients: [to("committee@agency.gov")] },                       // to the group address
  { id: "s2", conversationId: "c2", sentDateTime: "2026-07-02T10:00:00Z",
    toRecipients: [to("ann@example.com")], ccRecipients: [to("bob@example.com"), to("carol@agency.gov")] }, // 3 members
  { id: "s3", conversationId: "c3", sentDateTime: "2026-07-03T10:00:00Z",
    toRecipients: [to("ann@example.com")] },                            // 1:1, not a group send
];

check("group sends: group address + large send, not 1:1",
  A.groupSends(sent, { groupAddress: "Committee@Agency.gov", rosterEmails: roster, largeSendMin: 3 })
    .map(function (m) { return m.id; }),
  ["s1", "s2"]);

check("group sends: threshold respected",
  A.groupSends(sent, { groupAddress: "", rosterEmails: roster, largeSendMin: 4 }).length,
  0);

/* ------------------------------------------------------------ participation */

var fullRoster = [
  { id: "u1", displayName: "Ann Lee", mail: "ann@example.com" },
  { id: "u2", displayName: "Bob Ray", mail: "bob@example.com" },
  { id: "u3", displayName: "Carol Fox", mail: "carol@agency.gov" },
  { id: "u4", displayName: "Dan Poe", mail: "dan@example.com" },
];

var rows = A.buildParticipation({
  roster: fullRoster,
  ndrs: [{ subject: "Undeliverable: agenda", receivedDateTime: "2026-07-10T08:00:00Z",
    body: { content: "delivery failed: dan@example.com" } }],
  sent: sent,
  inbound: [
    // Ann replied to the group send on c1 AFTER it went out -> reply signal
    { conversationId: "c1", receivedDateTime: "2026-07-01T15:00:00Z",
      from: { emailAddress: { address: "ann@example.com" } } },
    // Bob's message on an unrelated conversation -> no signal
    { conversationId: "zz", receivedDateTime: "2026-07-05T15:00:00Z",
      from: { emailAddress: { address: "bob@example.com" } } },
  ],
  meetings: [
    { date: "2026-06-20T14:00:00Z", subject: "Monthly sync", attendees: ["carol@agency.gov"] },
  ],
  manual: { "carol@agency.gov": ["2026-07-15"] },
  opts: { groupAddress: "committee@agency.gov", largeSendMin: 3 },
});

check("bounced sorts first", rows[0].email, "dan@example.com");
check("bounced status", rows[0].status, "bounced");
check("unknown second (Bob: unrelated reply is no signal)", rows[1].email, "bob@example.com");
check("unknown status", rows[1].status, "unknown");
check("active sorted stalest first (Ann reply 7/1 before Carol manual 7/15)",
  rows.slice(2).map(function (r) { return r.email; }),
  ["ann@example.com", "carol@agency.gov"]);
check("Ann's signal is reply", rows[2].lastSeen.signal, "reply");
check("Carol: manual 7/15 beats teams 6/20", rows[3].lastSeen.signal, "manual");
check("Carol lastSeen date", rows[3].lastSeen.date, "2026-07-15");

// a reply BEFORE the group send is not a reply to it
var rowsEarly = A.buildParticipation({
  roster: [fullRoster[0]],
  ndrs: [], meetings: [], manual: {},
  sent: [sent[0]],
  inbound: [{ conversationId: "c1", receivedDateTime: "2026-06-30T09:00:00Z",
    from: { emailAddress: { address: "ann@example.com" } } }],
  opts: { groupAddress: "committee@agency.gov", largeSendMin: 3 },
});
check("reply before send ignored", rowsEarly[0].status, "unknown");

// bounce wins even when other signals exist
var rowsBounce = A.buildParticipation({
  roster: [fullRoster[0]],
  ndrs: [{ subject: "Undeliverable: z", receivedDateTime: "2026-07-20T08:00:00Z",
    body: { content: "ann@example.com failed" } }],
  sent: [], inbound: [],
  meetings: [{ date: "2026-07-18T14:00:00Z", attendees: ["ann@example.com"] }],
  manual: {},
  opts: { groupAddress: "", largeSendMin: 3 },
});
check("bounce outranks attendance", rowsBounce[0].status, "bounced");
check("bounce keeps lastSeen evidence", rowsBounce[0].lastSeen.signal, "teams");

// member with no mail (nested group / odd object) survives with UPN fallback
var rowsUpn = A.buildParticipation({
  roster: [{ id: "u9", displayName: "UPN Only", userPrincipalName: "upn@x.gov" }],
  ndrs: [], sent: [], inbound: [], meetings: [], manual: { "upn@x.gov": ["2026-07-01"] },
  opts: { groupAddress: "", largeSendMin: 3 },
});
check("UPN fallback works", rowsUpn[0].status, "active");

/* -------------------------------------------------------------------- CSV */

var csv = A.buildRemovalCsv([rows[0], rows[1]], 'Committee, "Main"');
var lines = csv.trim().split("\r\n");
check("csv header",
  lines[0],
  "email,name,status,last_seen,signal,bounce_date,bounce_subject,exchange_powershell");
check("csv bounced row core",
  lines[1].indexOf("dan@example.com,Dan Poe,bounced,never seen,,2026-07-10") === 0,
  true);
check("csv PowerShell line quoted (group name has comma+quotes)",
  lines[1].indexOf("Remove-DistributionGroupMember -Identity 'Committee, \"\"Main\"\"'") !== -1,
  true);
check("csv unknown row shows never seen",
  lines[2].indexOf("bob@example.com,Bob Ray,unknown,never seen") === 0,
  true);

check("csvEscape comma", A._internals.csvEscape("a,b"), '"a,b"');
check("csvEscape quote", A._internals.csvEscape('say "hi"'), '"say ""hi"""');

// CSV formulas and PowerShell expansion must remain literal user text.
check("csvEscape formula", A._internals.csvEscape("=1+1"), "'=1+1");
check("csvEscape whitespace formula", A._internals.csvEscape("\t@SUM(1)"), "'\t@SUM(1)");
check("csvEscape CR", A._internals.csvEscape("a\rb"), '"a\rb"');
var literal = A.buildRemovalCsv([rows[0]], "O'Hare $(Get-Process) `$HOME");
check("PowerShell arguments use literal quoting", literal.indexOf("-Identity 'O''Hare $(Get-Process) `$HOME' -Member 'dan@example.com'") !== -1, true);

if (failures) {
  console.error("\n" + failures + " auditor test(s) FAILED");
  process.exit(1);
}
console.log("All auditor tests passed.");
