# Mac Mail Helper Implementation Plan

Goal: Read existing Mac Mail accounts into the web directory without Entra.
Architecture: loopback-only Python server + read-only JXA + bundled browser UI + native launcher.
Spec: ../specs/2026-11-08-mac-mail-helper.md

- [ ] Test bridge request authorization, fixed commands, failure handling; implement server and Mail reader.
- [ ] Test local scanner adapter batch behavior and folder atomicity; integrate existing UI.
- [ ] Build Mac app with bundled web assets, document Python/macOS requirements.
- [ ] Run automated checks, inspect packaged UI and request user-handled macOS permission if needed.
- [ ] Review changes, push website and publish helper; verify live URLs.
