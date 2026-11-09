# Mac Mail local helper

User selected the local Mac Mail helper instead of Microsoft application registration or mailbox export. Existing Mail accounts and nested folders are read using Mail's installed scripting dictionary. No account password or Microsoft Client ID required.

The helper runs only on loopback and serves the same bundled browser directory UI. This avoids browser HTTPS/private-network restrictions. A random per-launch secret is passed in a fragment, removed immediately and kept in memory. All API calls require the secret, exact local Origin and Host. No external CORS. Only fixed read operations are available. No mail modification, sending or persistent mailbox cache.

Read folder snapshots by message IDs, then fetch bounded batches; error if IDs disappear rather than silently declare a full scan. Folder results commit atomically through the existing scanner. All local folders are rescanned on each run (deduplication remains); no misleading delta claim. Optional signature extraction is slower. macOS automation permission must be granted by the user. Only messages available through Mail can be read.

A native Mac app launcher presents Open Directory and Quit, runs a bundled dependency-free Python server, and opens local UI. Python 3 is required (present on this Mac). Distribute app as a same-origin GitHub Pages ZIP and local outputs. Apple Silicon package; Intel cross-compilation unavailable in installed toolchain. Unsigned app may require user approval in macOS. Hosted web page offers helper download; Microsoft setup remains optional.
