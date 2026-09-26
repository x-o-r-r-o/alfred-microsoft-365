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

## Status
Implemented and tested against a local mock of the Microsoft identity platform and Graph
(`python3 tests/test_m365.py`). Not yet verified against a live tenant; needs the user's own
Entra ID app registration (README ## Setup). Keyword for sign-in is `m365` (the Gallery requires ≥ 3 characters).

## Tech
- **Stack:** zsh + JXA; Microsoft Graph with device-code OAuth.
- **Dependencies:** Work/school tenants may need admin consent.
- Output via Alfred Script Filter JSON; settings via Workflow Configuration (`userconfigurationconfig`).
- Secrets (API keys/tokens) in the macOS Keychain, never in `prefs.plist`.
- Target: macOS 13+ on Apple Silicon and Intel.

## Milestones
1. Script filter prototype for the main keyword
2. Actions + modifiers, Universal Actions / File Actions where relevant
3. Workflow Configuration, icons, error states (no network / missing dependency)
4. README with screenshots, `tools/build.py --package` release, forum post, then Gallery submission when invited

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
