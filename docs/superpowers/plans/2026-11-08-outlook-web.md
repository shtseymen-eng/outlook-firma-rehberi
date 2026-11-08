# Outlook Web Implementation Plan

> Inline execution in the current chat. User selected Outlook Firma Rehberi as the first conversion.

Goal: Publish the working web interface and preserve the desktop program's reading/classification/export behavior.
Architecture: Static browser app, Microsoft MSAL PKCE, Graph GET-only client, session-memory directory and checkpoints, portable local backup, ExcelJS workbook.
Spec: docs/superpowers/specs/2026-11-08-outlook-web-design.md

- [x] Contact engine: normalize emails, extract unique sender signature only, classify, merge newest high-quality evidence, exclude own addresses. Unit tests with synthetic messages, quoted email and malicious-looking strings.
- [x] Graph scanning: trusted Graph URLs only, recursive paginated folders, throttling, cancellation and atomic folder checkpoints. Tests with mocked HTTP pages and mid-folder failures.
- [x] Excel and backup: five correctly formatted sheets, string cells, source attribution, incomplete scan flag, v4 import. Round-trip workbook tests.
- [x] Authentication and UI: multiple accounts, recursive folder selection, aliases, preview/editor, scan and cancellation, save/import, truthful unconfigured state. Build and browser QA.
- [x] Publish: GitHub source/build commit, Pages deploy, URL verification.
- [ ] Microsoft registration and real mailbox verification: pending explicit user approval and Microsoft account sign-in.

Verification: 17 Node tests passed; Vite production build passed; code review company-edit export finding fixed and regression-tested; synthetic v4 import checked in browser; GitHub Pages workflow 42466983321 completed successfully for 91aa7ce. Microsoft Client ID remains unconfigured and the live UI displays this accurately.
