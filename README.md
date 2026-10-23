# Roster Auditor

Outlook add-in that finds the ghosts on your group rosters. Pick a Microsoft
365 group or distribution list you belong to and it cross-references the
member list against four participation signals you already possess:

1. **Bounced sends (NDRs)** — recognized delivery failures to review,
   not proof that an address is dead. Later participation supersedes older failures.
2. **In-person attendance you record** — a one-click "Present ✓" grid for
   sign-in-sheet meetings, saved in your own mailbox settings.
3. **Teams meeting attendance** — Microsoft Graph attendance reports, which
   exist only for meetings *you organized* (by platform design — the add-in
   can never observe anyone else's meetings).
4. **Replies to your group sends** — the weakest signal, display-only:
   a member who reads without replying is not a ghost.

Each member gets a last-seen date with the evidence behind it. No signal at
all shows as **"engagement unknown"** — never "inactive." From there:

- **✉️ Re-engage** — drafts a polite "are you still an active member?" email
  to the members you select (the default action).
- **🗑 Remove** — two-click (arm, confirm) removal, available only on
  Microsoft 365 groups you own.
- **CSV export** — for classic distribution lists and mail-enabled security
  groups, which are **read-only via Microsoft Graph** (Exchange-mastered):
  a removal-recommendation CSV with member, status, evidence, and a
  ready-made `Remove-DistributionGroupMember` PowerShell line per row for
  your Exchange admin.

## Privacy posture

No backend, no accounts, no analytics, no AI. All Microsoft Graph access is
delegated — the add-in sees only what the signed-in user can already see —
and nothing leaves the Microsoft 365 boundary. A human confirms every action.
See `privacy-policy.html`.

## Architecture

Same shape as the other PurposeBuilt add-ins:

- `manifest.xml` — add-in-only XML manifest (task pane on the message-read
  ribbon; desktop/web, Mailbox 1.5).
- `src/auditor.js` — pure participation logic (group classification, NDR
  parsing, group-send detection, signal merge, CSV builder). No Office.js,
  no DOM — unit-tested under Node.
- `src/graph.js` — Microsoft Graph data layer; MSAL nested app auth (NAA),
  no backend. Delegated scopes: `User.Read`, `Group.Read.All`, `Group.ReadWrite.All`,
  `Mail.ReadWrite`, `Calendars.Read`, `OnlineMeetings.Read`,
  `OnlineMeetingArtifact.Read.All`.
  `User.Read` is required for the signed-in user's membership and ownership
  endpoints; the group scopes alone do not authorize `/me/memberOf`.
- `src/taskpane/` — the pane UI.
- Hosted on GitHub Pages; the manifest points at the hosted files.

## Development

```sh
npm test        # run the offline unit tests for src/auditor.js
npm run validate # validate manifest.xml (needs npm i first)
```

Contributions welcome — see [CONTRIBUTING.md](CONTRIBUTING.md).

## License

MIT © PurposeBuilt Systems
