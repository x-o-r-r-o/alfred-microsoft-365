# Microsoft 365 — Plan

**Priority tier:** 3 · **Bundle ID:** `com.xorro.microsoft-365`

## Why build it
Raycast demand this workflow replaces (downloads, 2026-09-26):

| Raycast extension | Downloads |
|---|---|
| Microsoft Teams | 20,150 |
| OneNote | 4,249 |
| **Total** | **24,399** |

**Alfred today:** 'Microsoft Teams' (2020) is minimal; Outlook workflows date from 2013–18; no OneNote.

## Features (v1.0)
- [ ] `teams` join next meeting, open chat with person, set status
- [ ] `onenote` search pages, create quick note in default section
- [ ] `outlook` search mail, today's calendar

## Tech
- **Stack:** zsh + JXA; Microsoft Graph with device-code OAuth.
- **Dependencies:** Work/school tenants may need admin consent.
- Output via Alfred Script Filter JSON; settings via Workflow Configuration (`userconfigurationconfig`).
- Secrets (API keys/tokens) in the macOS Keychain, never in `prefs.plist`.
- Target: macOS 13+ on Apple Silicon and Intel (universal binaries for any Swift helpers).

## Milestones
1. Script filter prototype for the main keyword
2. Actions + modifiers, Universal Actions / File Actions where relevant
3. Workflow Configuration, icons, error states (no network / missing dependency)
4. README with screenshots, `build.sh` release, submit to Alfred Gallery + forum post
