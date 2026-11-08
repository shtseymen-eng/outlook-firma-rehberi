# Outlook Web Implementation Plan

> Inline execution in the current chat. User selected Outlook Firma Rehberi as the first conversion.

Goal: Publish the working web interface and preserve the desktop program's reading/classification/export behavior.
Architecture: Static browser app, Microsoft MSAL PKCE, Graph GET-only client, session-memory directory and checkpoints, portable local backup, ExcelJS workbook.
Spec: docs/superpowers/specs/2026-11-08-outlook-web-design.md

- [ ] Contact engine: normalize emails, extract unique sender signature only, classify, merge newest high-quality evidence, exclude own addresses. Unit tests with synthetic messages, quoted email and malicious-looking strings.
- [ ] Graph scanning: trusted Graph URLs only, recursive paginated folders, throttling, cancellation and atomic folder checkpoints. Tests with mocked HTTP pages and mid-folder failures.
- [ ] Excel and backup: five correctly formatted sheets, string cells, source attribution, incomplete scan flag, v4 import. Round-trip workbook tests.
- [ ] Authentication and UI: multiple accounts, recursive folder selection, aliases, preview/editor, scan and cancellation, save/import, truthful unconfigured state. Build and browser QA.
- [ ] Publish: GitHub source/build commit, Pages deploy, URL verification. Microsoft registration changes require a concrete configuration and user approval at the final step.
