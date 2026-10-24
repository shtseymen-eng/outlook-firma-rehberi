# Outlook Firma Rehberi v5 Cross-Platform Design

## Intent

Upgrade the existing Windows-only Outlook Firma Rehberi repository into one downloadable desktop application for Windows and macOS. The application is for people who need to turn the mail history of one or more Microsoft mailboxes into a deduplicated, company-grouped Excel directory without granting the application permission to send, delete, move, or edit mail.

The existing signature parsing, contact enrichment, corporate/personal split, manual editing, and Excel concepts remain recognizable. Windows-only Outlook COM access and the Tkinter UI are replaced so both operating systems run the same behavior from one codebase.

## Goals

- Ship Windows `.exe` and macOS `.app`/`.dmg` artifacts from the same GitHub repository and release.
- Let a user add one or more Microsoft 365, Outlook.com, or Hotmail accounts through Microsoft sign-in.
- List every connected account and its complete folder tree, including hidden folders and descendants at any depth.
- Let the user select one or more accounts or folders and scan all historical messages in those selections.
- Read incoming, sent, deleted, archive, custom, nested, and deeply nested folders.
- Collect unique external addresses from sender, To, CC, and available BCC participants; exclude the connected user's own addresses.
- Enrich sender contacts from the signature while avoiding attribution of the sender's signature to unrelated recipients.
- Deduplicate globally by normalized email address and merge richer information found in later messages.
- Export one `.xlsx` file with category sheets, deterministic company grouping, and vertically merged company cells.
- Keep credentials and extracted data on the user's computer.
- Show scan progress, completion, and a native Save As dialog with editable filename and destination.

## Non-Goals

- Sending, deleting, moving, editing, marking, or otherwise changing any mailbox content.
- Connecting directly to Gmail, Yahoo, or Yandex mailboxes. Those domains are classified when found as contacts inside a connected Microsoft mailbox.
- Uploading message content or the generated directory to a project-operated server.
- A mobile, browser-only, or server-hosted version.
- Automatic access to accounts merely because they are open in the Outlook desktop application.

## Anti-Goals

- Do not request `Mail.ReadWrite`, `Mail.Send`, application-wide mailbox access, or tenant administrator credentials.
- Do not store full message bodies after parsing.
- Do not produce duplicate email addresses across sheets.
- Do not use separate implementations with diverging Windows and macOS behavior.
- Do not silently skip failed accounts, folders, pages, or rate-limit errors.

## Constraints

- Microsoft Graph delegated `Mail.Read` is required because signature extraction needs message bodies. `User.Read` and standard OpenID/offline scopes are used for identity and session renewal.
- The Entra application registration must support organizational directories and personal Microsoft accounts as a public desktop client.
- Some corporate tenants can require an administrator to approve delegated mail access. The application must explain that condition instead of treating it as a password failure.
- Graph returns top-level folders separately from their children. Folder discovery must recursively traverse every child collection and follow every pagination link.
- Message listing is paginated and can be throttled. Requests must honor retry guidance and remain cancellable.
- A polished macOS download requires Apple signing and notarization credentials; a polished Windows download benefits from a code-signing certificate. CI supports secrets when supplied and produces clearly labelled unsigned artifacts otherwise.
- Version 5.0 supports Windows 10/11 x64 and macOS 12 or newer. macOS release jobs produce separate Apple Silicon and Intel downloads unless the build dependencies prove a universal binary reliable.
- The repository owner supplies and owns the production Entra application registration. Development builds read a local client ID; release builds read a GitHub repository variable. A build without a client ID may run the UI and tests but cannot sign in.

## Considered Approaches

### Selected: Python, PySide6, MSAL, Microsoft Graph, OpenPyXL

Retain the project's Python parsing and Excel investment, replace Tkinter with a modern cross-platform Qt UI, and replace COM with Graph. This creates one behavior on both operating systems, makes the existing business rules reusable, and keeps the rewrite focused on platform boundaries.

### Rejected: Electron and JavaScript rewrite

Electron can provide a polished UI, but it would require rewriting the mature Python parsing and Excel logic and would create much larger installers. It does not improve the core mailbox or export behavior enough to justify the migration.

### Rejected: Continue COM on Windows and add a separate macOS adapter

This would preserve the current Windows path but create two account models, two folder implementations, and inconsistent results. Outlook for macOS cannot provide Windows COM behavior, so the maintenance and test burden would remain permanently split.

## Application Architecture

The repository becomes a package with focused modules instead of adding more responsibilities to the current 1,100-line `uygulama.py`:

- `outlook_rehberi/auth.py`: MSAL public-client sign-in, account selection, token renewal, and secure cache integration.
- `outlook_rehberi/graph.py`: Graph requests, pagination, retry/backoff, recursive folder discovery, message iteration, and cancellation.
- `outlook_rehberi/models.py`: account, folder, message participant, contact, scan statistics, and category data models.
- `outlook_rehberi/parser.py`: HTML-to-text, signature extraction, name/title/company/phone parsing, and participant conversion.
- `outlook_rehberi/classifier.py`: personal-domain, Polisan/Poliport, Gümrük/YGM, other-corporate classification and global deduplication.
- `outlook_rehberi/storage.py`: local settings, secure token-cache handoff, scan state, contact cache, and incremental checkpoints.
- `outlook_rehberi/exporter.py`: workbook layout, ordering, styling, merged company cells, and validation.
- `outlook_rehberi/ui/`: PySide6 windows, account/folder tree, progress screen, directory preview, completion dialog, and Save As flow.
- `main.py`: application entry point and dependency assembly only.

`uygulama.py` remains temporarily as the v4 reference during migration, then becomes a thin compatibility launcher or is retired once v5 feature parity is proven.

## Account and Permission Flow

1. On first launch, the application shows an empty connected-account list and an Add Account action.
2. Add Account opens the system browser for Microsoft sign-in and consent.
3. After consent, only refreshable delegated tokens are retained through the operating system's secure credential store: macOS Keychain or Windows Credential Manager.
4. Each additional account is added through the same explicit flow. The application never borrows Outlook desktop sessions.
5. The main screen lists connected accounts. Removing an account deletes its local session and cached scan state without touching the mailbox.
6. If a tenant blocks consent, the UI displays the Microsoft-provided reason and tells the user that their administrator must approve read access.

## Folder and Scan Flow

1. The user selects one or more connected accounts.
2. The application loads each account's root folders and recursively loads all child folders, including hidden folders.
3. A tri-state tree lets the user select whole accounts, folders, or nested subtrees. Parent selection includes every descendant.
4. Start Scan creates one queued scan job covering every selected folder without scanning a descendant twice.
5. The first scan walks all message pages for all historical dates. Later scans use saved per-folder checkpoints/delta links when available; the user can request a full rebuild.
6. Incoming messages contribute the sender and other external participants. Sent messages contribute recipients. Selected account addresses and obvious system/no-reply addresses are excluded from the directory.
7. Signature-derived name, title, company, and phone data are applied only to the sender represented by that signature. Recipient addresses are still recorded with whatever reliable information Graph exposes.
8. Progress shows current account, folder, processed-message count, unique-address count, warnings, and a Cancel action.
9. Partial failures are summarized. Successful data is retained, but export identifies incomplete accounts/folders until the user accepts or retries.
10. On success the application displays “Tarama tamamlandı” and opens the operating system's Save As dialog. The default name is `Outlook_Firma_Rehberi_YYYY-MM-DD_HH-mm.xlsx`; the user can change name and destination.

The default contact cache is an “ever observed” directory: removing a source message, deselecting a folder, or disconnecting an account does not silently delete a previously discovered contact. Full Rebuild clears the directory and recomputes it only from the accounts and folders selected for that rebuild. A folder checkpoint advances only after that folder completes successfully. Cancellation does not advance the active folder. An incomplete scan does not open Save As automatically; the user may retry or explicitly export a workbook whose summary is marked incomplete.

Version 5.0 traverses the ordinary mailbox folder tree of each explicitly connected Microsoft account. Microsoft 365 Online Archive/In-Place Archive mailboxes and shared/delegated mailboxes are outside this release because they require a separate mailbox and permission model.

## Contact Normalization and Deduplication

- Lowercase and trim email addresses; normalize Unicode domains to a stable comparison form.
- Use normalized email as the global unique key across all accounts, folders, and workbook sheets.
- Merge non-empty names, titles, company labels, mobile numbers, landlines, source accounts, and source folders across occurrences.
- Prefer signature-derived values over Graph display names, and Graph display names over domain-derived guesses. For equally reliable conflicting values, choose the value from the newest message timestamp; if timestamps are equal, use normalized lexical order. Phone values and source values are deduplicated unions sorted in a stable order.
- Never duplicate a contact in more than one category. Classification precedence is:
  1. `polisan.com.tr` or `poliport.com`
  2. known personal domain
  3. Gümrük/YGM keyword match
  4. other corporate domain
- Personal domains include the existing list and common variants of Gmail, Hotmail, Outlook, Live, Yahoo, Yandex, iCloud, Proton, AOL, and MSN.
- Gümrük/YGM matching is case-insensitive and Turkish-character tolerant across normalized email/domain, signature company, title, and signature text. Initial terms include `gümrük`, `gumruk`, `gümrükleme`, `ygm`, `müşavirlik`, `musavirlik`, `customs`, and `customs brokerage`.
- A company block is keyed by normalized email domain. Different domains are not merged automatically even if their extracted labels resemble one another; this avoids false company merges. The selected best company label is displayed for that domain. Polisan and Poliport are the explicit exception: they share one worksheet but remain two domain blocks.
- A connected account's Graph `mail` and `userPrincipalName` values are always treated as the user's own addresses. Account settings also provide an optional local alias list so proxy aliases can be excluded without requesting broader directory permissions.
- Signature parsing examines only the current message's unique content before common quoted-reply and forwarded-message markers. Older quoted signatures are not attributed to the current sender.

## Excel Workbook Design

The export contains one workbook with these sheets in order:

1. `Özet`
2. `Polisan-Poliport`
3. `Gümrük-YGM`
4. `Diğer Kurumsal`
5. `Kişisel Mailler`

Every detail sheet uses the same columns:

| Column | Value |
| --- | --- |
| A | Ad Soyad |
| B | Görev / Unvan |
| C | Domain |
| D | Firma |
| E | E-posta |
| F | Cep Telefonu |
| G | Sabit Telefon |
| H | Kaynak Hesap(lar) |
| I | Kaynak Klasör(ler) |

Rows are sorted by normalized company, domain, then email. Contacts from the same company/domain form one contiguous block. Column D is vertically merged over the block, matching the requested relationship between the company cell and adjacent email rows in column E. Column C is also merged when the block has one domain. Filters, frozen headers, alternating company-block colors, and sensible widths remain enabled.

`Polisan-Poliport` contains two blocks in one sheet, treating `polisan.com.tr` and `poliport.com` as departments of the same business group. The summary contains per-category totals, globally unique total contacts, scanned account/folder counts, message count, completion time, and any incomplete-scan warning.

## Local Data and Privacy

- Store application settings and normalized contacts in the operating system's application-data directory, never beside the executable.
- Store the complete serialized MSAL token cache through `keyring`, backed by Keychain on macOS and Credential Manager on Windows. This includes the access, ID, refresh, and account metadata MSAL requires for reliable silent multi-account sign-in; it remains encrypted by the operating system credential store.
- Do not persist complete message bodies. Keep only extracted contact fields, source metadata, folder checkpoints, and minimal diagnostic information.
- Provide Disconnect Account, Clear Scan Data, and Full Rebuild actions.
- Logs redact access tokens, message bodies, and email local-parts where practical.

## Error Handling

- Retry transient network failures and Graph throttling with bounded exponential backoff and server-provided retry delays.
- Follow pagination links exactly and reject loops.
- Record and surface permission-denied, tenant-consent, expired-session, unavailable-mailbox, and deleted-folder conditions per account.
- Continue past malformed individual messages while counting and reporting them.
- Write Excel to a temporary file in the selected directory, validate it, then atomically replace the destination so interrupted exports do not leave corrupt files.
- Keep the UI responsive by running network and export work outside the UI thread and delivering state changes through Qt signals.

## Packaging and Release

- PyInstaller builds platform-specific executables from the common source.
- GitHub Actions runs unit/integration tests on Windows and macOS, then produces `.exe`, `.app`, and `.dmg` artifacts.
- A tag release attaches the Windows x64 installer plus Apple Silicon and Intel macOS DMG downloads, installation notes, and checksum files.
- CI signs and notarizes when the repository has the required Apple and Windows signing secrets. Unsigned builds remain available for development and are clearly labelled.
- The README explains Microsoft account support, read-only permissions, first sign-in, multiple accounts, folder selection, Excel categories, and platform security prompts.

## Testing and Acceptance

- Unit tests cover address normalization, global deduplication, contact enrichment, category precedence, Turkish keyword matching, and personal-domain detection.
- Workbook tests load the generated file and verify sheet order, global uniqueness, grouping order, merged D-column ranges, summary totals, and save-path behavior.
- Graph tests use recorded/mock responses for multi-page folders, unlimited nesting, message pagination, 429 retry, partial failure, cancellation, and multiple accounts.
- Authentication tests mock MSAL and secure storage; no live credentials enter the test suite.
- UI tests cover account/folder selection, disabled scan with no selection, progress state, completion dialog, cancelled Save As, and successful export.
- Packaging smoke tests launch the built application on both CI operating systems.

The feature is accepted when a user can install the appropriate download, add at least two Microsoft accounts, select nested folders from both, scan all history read-only, and save one Excel workbook in which each normalized address appears exactly once and every row belongs to the correct requested category.

