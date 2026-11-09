# Mac Mail Helper Implementation Plan

Goal: Read existing Mac Mail accounts into the web directory without Entra.
Architecture: loopback-only Python server + read-only JXA + bundled browser UI + native launcher.
Spec: ../specs/2026-11-08-mac-mail-helper.md

- [x] Test bridge request authorization, fixed commands, failure handling; implement server and Mail reader.
- [x] Test local scanner adapter batch behavior and folder atomicity; integrate existing UI.
- [x] Build Mac app with bundled web assets, document Python/macOS requirements.
- [ ] Run automated checks, inspect packaged UI and request user-handled macOS permission if needed.
- [ ] Review changes, push website and publish helper; verify live URLs.

Validation: 22 Node tests + 4 Python integration tests passed; packaged app extracted and codesign verified. Browser synthetic scan: 2 messages -> 1 contact. Real Mail automation permission not yet granted/tested. Added desktop download catalog with existing Outlook v4.6 Windows and Poliport v1.0 Windows/Mac packages. Existing Windows executables not executed on this Mac.
