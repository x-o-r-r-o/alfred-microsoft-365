# Microsoft 365 — Plan

**Priority tier:** 3 · **Bundle ID:** `io.github.x-o-r-r-o.microsoft-365` · **Keywords:** `teams`, `onenote`, `outlook`

## Why build it
Raycast demand this workflow replaces (downloads, 2026-09-26):

| Raycast extension | Downloads |
|---|---|
| Microsoft Teams | 20,150 |
| OneNote | 4,249 |
| **Total** | **24,399** |

**Alfred today:** 'Microsoft Teams' (2020) is minimal; Outlook workflows date from 2013–18; no OneNote.

## Features (v1.0)
- [x] `teams` join next meeting, open chat with person, set status
- [x] `onenote` search pages, create quick note in default section
- [x] `outlook` search mail, today's calendar

## Features (v1.1, round 4)
- [x] Teams status message: `teams message <text> [:: duration]`, shown and clearable from `teams message` / `teams status` (Presence.ReadWrite, work or school accounts only)
- [x] ⌃↩ on a person copies the Teams chat link (requested in raycast/extensions#16808)
- [x] Join links for meetings without `onlineMeeting` (other organizations' Teams invites, Zoom, Google Meet, Webex, GoTo, Whereby, Chime) from the location or the body preview; lookalike hosts are ignored

## Status
Implemented and tested against a local mock of the Microsoft identity platform and Graph
(`python3 tests/test_m365.py`). Not yet verified against a live tenant; needs the user's own
Entra ID app registration (README ## Setup). Keyword for sign-in is `m365` (the Gallery requires ≥ 3 characters).

## Known limitations
- Tested only against a local mock of the Microsoft identity platform and Graph; no live tenant yet.
- OneNote search matches page titles only (Graph has no full-text page search); the index holds the 5,000 most recently edited pages and refreshes every 30 minutes.
- Teams status needs a work or school account, and only shows while the user is signed in to a Teams app.
- Chat and call links use teams.microsoft.com; personal (teams.live.com) accounts may not open them.
- Rate limiting: Alfred's automatic queue delay plus "terminate previous script" (`queuemode` 2 on the `teams` and `outlook` filters since v1.1; v1.0.0 shipped `queuemode` 1, "wait"), a per-query cache and a shared back-off after 429 keep requests low; there is no request budget beyond that (Graph's per-app limits are far above keystroke rates).
- The device code poller receives the device code in its environment (visible to the same user's `ps -E` for up to 15 minutes); tokens themselves never leave the Keychain except on curl's stdin.

## Verify in real Alfred
- [ ] Device code sign-in with a work account and with a personal account (tenant `consumers`), including the notification from the background poller (External Trigger `notify`).
- [ ] Joining a meeting opens the Teams app; with “Web browser” selected it opens the browser.
- [ ] Setting and resetting the status shows in Teams; the 7-day cap is accepted by Graph.
- [ ] People search, chat and video-call links open the right chat.
- [ ] OneNote index on a large account (section-by-section fallback), `onenote:` links open the app, new pages land in the configured section.
- [ ] Mail search with `from:` / `subject:` syntax; the webLink opens the message.
- [ ] Holding ⌘/⌥ on rows without that modifier behaves sensibly (Outlook mail has no ⌥ action).
- [ ] Status message with and without `:: 2h`; the message and its expiry show in Teams; clearing works.
- [ ] ⌃↩ on a person copies a chat link that opens the chat.
- [ ] A Zoom/Meet invitation and a Teams invitation from another organization show a working ⌘↩ join in `outlook`.
- [ ] Typing fast in `outlook` / `teams` terminates the previous search (queuemode 2) without leaving stale locks.
- [ ] No empty notification after ↩ on a meeting, page or mail (actions print nothing).
- [ ] Screenshots for every README paragraph (including `images/teams-message.png`).

## Tech
- **Stack:** zsh + JXA; Microsoft Graph with device-code OAuth.
- **Dependencies:** Work/school tenants may need admin consent.
- Output via Alfred Script Filter JSON; settings via Workflow Configuration (`userconfigurationconfig`).
- Secrets (API keys/tokens) in the macOS Keychain, never in `prefs.plist`.
- Target: macOS 13+ on Apple Silicon and Intel.

## Milestones
1. [x] Script filter prototype for the main keyword
2. [x] Actions + modifiers, Universal Actions / File Actions where relevant
3. [x] Workflow Configuration, icons, error states (no network / missing dependency)
4. [ ] README with screenshots, `tools/build.py --package` release, forum post, then Gallery submission when invited

## Release checklist (Alfred forum + Gallery)
Sources: alfred.app/submit, alfred.app/submit/styleguide, alfred.app/submit/screenshots, alfredforum.com topics 23976 and 23388.

- [x] README starts with `## Usage`; each paragraph ends "via the `kw` keyword" / "via the Universal Action"
- [ ] A clean screenshot (window only, transparent background, real-looking data, no other workflows) after each paragraph, stored in `images/`
- [x] Modifiers listed as `* <kbd>⌘</kbd><kbd>↩</kbd> Action.`; Quick Look written as <kbd>⌘</kbd><kbd>Y</kbd>
- [x] `## Setup` only for genuine manual steps (no app installs or API keys; the Gallery lists those)
- [x] Every keyword is ≥ 3 characters and configurable via `{var:keyword_*}`
- [x] Settings in Workflow Configuration; the info.plist `readme` (About This Workflow) matches README.md
- [x] Main icon ≥ 256×256 px
- [x] No self-updater; never download or install software (no pip/brew/curl of binaries); dependencies declared for Alfred to handle
- [x] Any compiled binary is Developer ID signed + notarised; never strip quarantine (none: zsh/bash + JXA only)
- [x] No hard-coded paths; `prefs.plist` is git-ignored; secrets stay in Keychain
- [ ] AI assistance disclosed in the README and the forum post (README done; forum post pending)
- [ ] Version bumped in `workflow.json`; `python3 tools/build.py --package`; GitHub release with the `.alfredworkflow` attached
- [ ] Forum post in "Share your Workflows" with a screenshot, keywords, and the GitHub link

## Round 4 notes (post-release audit)
- `tools/build.py` now reads an optional `queuemode` per Script Filter (default 1, as before). Port this to the canonical copy in alfred-devtoolbox so the next sync doesn't drop it.
- Verified against official alfredapp plists: network Script Filters (google-drive, *-suggest) use `queuemode` 2 = terminate; local ones use 1 = wait.
- README Setup re-checked against Microsoft Learn (quickstart-register-app, updated 2026-06): the account-type drop-down now reads “Single tenant only / Multiple Entra ID tenants / Any Entra ID Tenant + Personal Microsoft accounts / Personal accounts only”, the registration form no longer asks for a redirect URI, and “Allow public client flows” is on the Authentication page's **Settings** tab.
- macOS 13: all JXA features (Unicode property escapes, NSTask `launchAndReturnError`, Security framework, NSDateFormatter templates) and curl config options used are available on macOS 13.0; no change needed.

## Ideas for v1.1
Ranked by value for effort (not implemented yet):
1. Recent chats list (`teams` with an empty query or `chat`): needs `Chat.Read`, one more consent; the Raycast extension's most used command is "Find Chat".
2. Tomorrow's agenda / next working day when today's meetings are over (`outlook tomorrow` would clash with mail search, so it needs a prefix like `agenda`).
3. A Hotkey (unset by default) that joins the meeting happening now or next.
4. Mail search ⌥↩: reply or forward on the web (`webLink` + `&action=reply` isn't documented; needs a live check).
5. Mark a mail as read from the result list (`Mail.ReadWrite`: a broader permission, so optional).
6. OneNote: full-text search is still impossible through Graph (`$search` on pages was beta-only and removed); a local index of page content would need `Notes.Read` downloads of every page, too slow and heavy for 5,000 pages. Users ask for it (raycast/extensions#23268).
7. Show the other person's presence next to people results (`Presence.Read.All`, which some tenants restrict to admin consent).
8. Out-of-office auto-reply on/off (`MailboxSettings.ReadWrite`).

