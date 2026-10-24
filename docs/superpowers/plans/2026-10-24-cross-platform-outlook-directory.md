# Outlook Firma Rehberi v5 Cross-Platform Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Upgrade the existing Windows-only application into one read-only Microsoft Graph desktop application that builds the requested deduplicated Excel directory on Windows and macOS.

**Architecture:** A focused Python package replaces the Windows COM and monolithic Tkinter boundaries while reusing the proven parsing concepts. PySide6 provides the shared desktop UI, MSAL and Microsoft Graph provide delegated multi-account mailbox access, SQLite and OS credential storage keep state local, and OpenPyXL produces the categorized workbook.

**Tech Stack:** Python 3.11+, PySide6 6.x, MSAL 1.x, Requests 2.x, keyring 25.x, platformdirs 4.x, OpenPyXL 3.1+, SQLite, pytest, pytest-qt, responses, PyInstaller, Inno Setup, GitHub Actions.

**Spec:** `docs/superpowers/specs/2026-10-24-cross-platform-outlook-directory-design.md`

## Global Constraints

- Support Windows 10/11 x64 and macOS 12+ with separate Apple Silicon and Intel downloads.
- Use Microsoft Graph delegated `Mail.Read`; never request `Mail.ReadWrite`, `Mail.Send`, or application mailbox permissions.
- Support Microsoft 365, Outlook.com, and Hotmail source accounts; Gmail/Yahoo/Yandex are contact classifications, not source mailbox providers.
- Keep full message bodies in memory only while parsing; never persist them.
- Normalize and deduplicate email addresses globally across every account, folder, and worksheet.
- Traverse folder trees and paginated collections without a fixed nesting-depth limit.
- Keep the UI responsive and every scan cancellable.
- Store the serialized MSAL cache only through the OS credential store.
- Do not advance an incomplete folder checkpoint.
- Build one workbook with sheet order `Özet`, `Polisan-Poliport`, `Gümrük-YGM`, `Diğer Kurumsal`, `Kişisel Mailler`.
- Show “Tarama tamamlandı” before opening the native Save As dialog; do not auto-open Save As for cancelled/incomplete scans.
- Preserve v4 data through a one-time, non-destructive `rehber_data.json` import.

## Review Focus

- A folder tree that is both deeply nested and paginated must visit every folder exactly once: pinned by `test_recursive_folder_tree_follows_pages_without_duplicates` in Task 5.
- The same address observed in multiple accounts plus a configured own-address alias must deduplicate or be excluded correctly: pinned by `test_multi_account_evidence_is_global_and_own_alias_is_excluded` in Task 7.
- Cancellation or a 429/error midway through a folder must not advance that folder's checkpoint: pinned by `test_incomplete_folder_does_not_commit_checkpoint` in Task 7.
- A quoted older signature must not be attributed to the current sender: pinned by `test_quoted_reply_signature_is_excluded` in Task 2.
- Workbook category uniqueness, D-column company merges, and atomic destination replacement must remain correct together: pinned by `test_workbook_is_globally_unique_grouped_merged_and_atomic` in Task 4.

---

### Task 1: Package Foundation, Configuration, and Domain Models

**Files:**
- Create: `pyproject.toml`
- Create: `outlook_rehberi/__init__.py`
- Create: `outlook_rehberi/__main__.py`
- Create: `outlook_rehberi/config.py`
- Create: `outlook_rehberi/models.py`
- Create: `tests/test_config.py`
- Create: `tests/test_models.py`
- Modify: `.gitignore`

**Interfaces:**
- Produces: `AppConfig.load(env: Mapping[str, str], resource_path: Path | None = None) -> AppConfig`
- Produces: `Category` enum values `POLISAN_POLIPORT`, `CUSTOMS_YGM`, `CORPORATE`, `PERSONAL`
- Produces: immutable `AccountRef`, `FolderRef`, `Participant`, `GraphMessage`, `GraphMessagePage`, `SourceRef`, `ContactEvidence`, `Contact`, `ScanSelection`, `ScanProgress`, `ScanWarning`, `ScanStats`, `ScanResult`, and `LegacyImportResult` dataclasses.
- Produces: `AccessTokenProvider.get_access_token(account_id: str) -> str` protocol.
- Produces: `CancellationToken.cancel() -> None` and `CancellationToken.raise_if_cancelled() -> None`.

- [ ] **Step 1: Write failing model and configuration tests**

Add tests asserting exact category values, immutable identity fields, cancellation behavior, `OUTLOOK_REHBERI_CLIENT_ID` precedence over resource configuration, and a user-facing configuration error when no client ID exists.

- [ ] **Step 2: Run the foundation tests and verify failure**

Run: `python -m pytest tests/test_config.py tests/test_models.py -v`
Expected: FAIL because the package and interfaces do not exist.

- [ ] **Step 3: Add project metadata and dependencies**

Define Python `>=3.11`, runtime dependencies, optional test dependencies, console entry point `outlook-firma-rehberi = outlook_rehberi.__main__:main`, and pytest configuration including `QT_QPA_PLATFORM=offscreen` for CI.

- [ ] **Step 4: Implement the minimal configuration and model interfaces**

Keep models free of PySide6, requests, MSAL, OpenPyXL, and filesystem side effects so every later task can test them in isolation.

- [ ] **Step 5: Run the foundation tests**

Run: `python -m pytest tests/test_config.py tests/test_models.py -v`
Expected: PASS.

- [ ] **Step 6: Commit the foundation**

Run: `git add pyproject.toml outlook_rehberi tests .gitignore && git commit -m "feat: add cross-platform v5 package foundation"`

### Task 2: Graph-Message Parsing and Safe Signature Extraction

**Files:**
- Create: `outlook_rehberi/parser.py`
- Create: `tests/fixtures/messages.py`
- Create: `tests/test_parser.py`
- Reference: `uygulama.py:50-220`

**Interfaces:**
- Consumes: `GraphMessage`, `Participant`, `SourceRef`, and `ContactEvidence` from Task 1.
- Produces: `html_to_text(html: str) -> str`
- Produces: `unique_message_content(text: str) -> str`
- Produces: `extract_signature(text: str) -> str`
- Produces: `sender_evidence(message: GraphMessage, source: SourceRef) -> ContactEvidence | None`
- Produces: `participant_evidence(participant: Participant, source: SourceRef, observed_at: datetime) -> ContactEvidence | None`
- Produces: `message_evidence(message: GraphMessage, source: SourceRef, own_addresses: AbstractSet[str]) -> list[ContactEvidence]`

- [ ] **Step 1: Write failing parser tests**

Cover Turkish and English signatures, HTML conversion, mobile/landline extraction, sender-only signature enrichment, sent-message recipients, To/CC/BCC participants, invalid addresses, own-address exclusion, and `test_quoted_reply_signature_is_excluded`.

- [ ] **Step 2: Run parser tests and verify failure**

Run: `python -m pytest tests/test_parser.py -v`
Expected: FAIL because `outlook_rehberi.parser` does not exist.

- [ ] **Step 3: Port and isolate the existing parser**

Move the proven regular expressions and signature heuristics out of COM objects. Accept only Task 1 models, stop current content at standard reply/forward headers, and attach quality ranks plus observation timestamps to inferred values.

- [ ] **Step 4: Run parser tests**

Run: `python -m pytest tests/test_parser.py -v`
Expected: PASS, including quoted-reply and recipient-attribution cases.

- [ ] **Step 5: Commit the parser**

Run: `git add outlook_rehberi/parser.py tests/fixtures tests/test_parser.py && git commit -m "feat: parse Graph messages and signatures safely"`

### Task 3: Global Contact Directory and Category Rules

**Files:**
- Create: `outlook_rehberi/classifier.py`
- Create: `tests/test_classifier.py`
- Reference: `uygulama.py:333-344`

**Interfaces:**
- Consumes: `Category`, `Contact`, and `ContactEvidence` from Task 1.
- Produces: `normalize_email(value: str) -> str | None`
- Produces: `normalize_search_text(value: str) -> str`
- Produces: `classify_contact(contact: Contact) -> Category`
- Produces: `company_group_key(contact: Contact) -> tuple[str, str]`
- Produces: `ContactDirectory.upsert(evidence: ContactEvidence) -> Contact`
- Produces: `ContactDirectory.contacts() -> tuple[Contact, ...]`
- Produces: deterministic serialization helpers for Task 6 storage.

- [ ] **Step 1: Write failing classification and merge tests**

Assert global lowercase/Unicode-domain uniqueness, stable phone unions, quality/newest/tie-break precedence, Polisan/Poliport priority, personal-domain priority over Gümrük keywords, Turkish-tolerant Gümrük/YGM matching, other-corporate fallback, and domain-based company blocks.

- [ ] **Step 2: Run classifier tests and verify failure**

Run: `python -m pytest tests/test_classifier.py -v`
Expected: FAIL because the classifier does not exist.

- [ ] **Step 3: Implement deterministic directory behavior**

Use normalized email as the only identity key. Preserve every source account/folder, choose scalar values by `(quality_rank, observed_at, normalized_lexical_value)`, and classify each final contact exactly once.

- [ ] **Step 4: Run classifier tests**

Run: `python -m pytest tests/test_classifier.py -v`
Expected: PASS with stable results under reversed evidence order.

- [ ] **Step 5: Commit classification**

Run: `git add outlook_rehberi/classifier.py tests/test_classifier.py && git commit -m "feat: deduplicate and classify directory contacts"`

### Task 4: Categorized Excel Workbook and Atomic Save

**Files:**
- Create: `outlook_rehberi/exporter.py`
- Create: `tests/test_exporter.py`
- Reference: `uygulama.py:346-427`

**Interfaces:**
- Consumes: `ContactDirectory`, `Category`, and `ScanResult` from Tasks 1 and 3.
- Produces: `build_workbook(contacts: Sequence[Contact], result: ScanResult) -> openpyxl.Workbook`
- Produces: `save_workbook_atomic(workbook: Workbook, destination: Path) -> Path`
- Produces: `default_export_name(now: datetime) -> str`

- [ ] **Step 1: Write failing workbook tests**

Assert exact sheet order and Turkish labels, columns A-I, email in E, company in D, stable company/domain/email sorting, separate Polisan and Poliport blocks on one sheet, summary totals, incomplete warning, and valid default filename.

- [ ] **Step 2: Add the combined export safety test**

Implement `test_workbook_is_globally_unique_grouped_merged_and_atomic`: load the saved workbook, assert every normalized email occurs once across all detail sheets, assert expected D-column merge ranges, and prove a failed temporary save does not replace an existing destination.

- [ ] **Step 3: Run exporter tests and verify failure**

Run: `python -m pytest tests/test_exporter.py -v`
Expected: FAIL because the exporter does not exist.

- [ ] **Step 4: Implement workbook generation and atomic replacement**

Write to a same-directory temporary `.xlsx`, reopen it with OpenPyXL for validation, then use `os.replace`. Merge D for every multi-row company block and C only when that block has one domain; do not merge single rows.

- [ ] **Step 5: Run exporter tests**

Run: `python -m pytest tests/test_exporter.py -v`
Expected: PASS.

- [ ] **Step 6: Commit the exporter**

Run: `git add outlook_rehberi/exporter.py tests/test_exporter.py && git commit -m "feat: export categorized company-grouped workbook"`

### Task 5: Read-Only Microsoft Graph Client

**Files:**
- Create: `outlook_rehberi/graph.py`
- Create: `tests/fixtures/graph_responses.py`
- Create: `tests/test_graph.py`

**Interfaces:**
- Consumes: `AccountRef`, `FolderRef`, `GraphMessage`, `CancellationToken` from Task 1.
- Consumes: `AccessTokenProvider.get_access_token(account_id: str) -> str` protocol from Task 1; Task 6 supplies its production implementation.
- Produces: `GraphClient.get_profile(account_id: str) -> AccountRef`
- Produces: `GraphClient.list_folder_tree(account: AccountRef, cancellation: CancellationToken) -> tuple[FolderRef, ...]`
- Produces: `GraphClient.iter_messages(account: AccountRef, folder: FolderRef, cancellation: CancellationToken, delta_link: str | None = None) -> Iterator[GraphMessagePage]`
- Produces: typed `GraphPermissionError`, `GraphThrottleError`, `GraphUnavailableError`, and `GraphProtocolError`.

- [ ] **Step 1: Write failing request, pagination, and retry tests**

Use `responses` to pin authorization headers, `$select` fields, hidden-folder requests, `@odata.nextLink`, bounded 429 `Retry-After`, 5xx retry, cancellation, and redacted error messages.

- [ ] **Step 2: Add recursive-tree and message-page tests**

Implement `test_recursive_folder_tree_follows_pages_without_duplicates` with at least four nesting levels and paginated children. Add message tests that verify body, sender, To, CC, BCC, dates, and delta/checkpoint metadata are mapped without persisting raw responses.

- [ ] **Step 3: Run Graph tests and verify failure**

Run: `python -m pytest tests/test_graph.py -v`
Expected: FAIL because the Graph client does not exist.

- [ ] **Step 4: Implement the minimal REST client**

Use one `requests.Session`, exact Graph-provided pagination URLs, a visited-URL/folder guard, bounded retry, and cancellation checks before requests and between yielded pages. Do not add write methods.

- [ ] **Step 5: Run Graph tests**

Run: `python -m pytest tests/test_graph.py -v`
Expected: PASS.

- [ ] **Step 6: Commit Graph access**

Run: `git add outlook_rehberi/graph.py tests/fixtures/graph_responses.py tests/test_graph.py && git commit -m "feat: add recursive read-only Graph mailbox client"`

### Task 6: Secure Multi-Account Authentication and Local Storage

**Files:**
- Create: `outlook_rehberi/auth.py`
- Create: `outlook_rehberi/storage.py`
- Create: `tests/test_auth.py`
- Create: `tests/test_storage.py`

**Interfaces:**
- Consumes: `AppConfig`, account/contact/checkpoint models, and classifier serialization from Tasks 1 and 3.
- Produces: `KeyringTokenCache.load() -> SerializableTokenCache`
- Produces: `KeyringTokenCache.save(cache: SerializableTokenCache) -> None`
- Produces: `MicrosoftAuthenticator.add_account() -> AccountRef`
- Produces: `MicrosoftAuthenticator.accounts() -> tuple[AccountRef, ...]`
- Produces: `MicrosoftAuthenticator.get_access_token(account_id: str) -> str`
- Produces: `MicrosoftAuthenticator.disconnect(account_id: str) -> None`
- Produces: `LocalStore.load_directory() -> ContactDirectory`
- Produces: `LocalStore.commit_folder(account_id: str, folder_id: str, delta_link: str | None, evidence: Sequence[ContactEvidence]) -> None`
- Produces: `LocalStore.aliases(account_id: str) -> frozenset[str]` and `set_aliases(...) -> None`
- Produces: `LocalStore.import_legacy_json(path: Path) -> LegacyImportResult`

- [ ] **Step 1: Write failing secure-cache and account tests**

Mock MSAL and keyring. Assert organizational/personal authority, exact read-only scopes, multiple cached accounts, silent renewal, browser login fallback, disconnect behavior, tenant-consent error mapping, and no token/cache text in logs or SQLite.

- [ ] **Step 2: Write failing storage and migration tests**

Use a temporary SQLite database. Assert atomic folder transactions, append-only default behavior, full-rebuild reset, aliases, schema migration, corrupt-row isolation, one-time legacy import, stable deduplication, and preservation of the original JSON file.

- [ ] **Step 3: Run auth/storage tests and verify failure**

Run: `python -m pytest tests/test_auth.py tests/test_storage.py -v`
Expected: FAIL because authentication and storage do not exist.

- [ ] **Step 4: Implement secure cache and authentication**

Persist the complete serialized MSAL cache as one keyring secret. Treat the public client ID as configuration, never a secret. Pass only the API scopes `Mail.Read` and `User.Read`; allow MSAL to manage its reserved OpenID and offline-access scopes.

- [ ] **Step 5: Implement SQLite storage and v4 migration**

Place the database under `platformdirs.user_data_dir("OutlookFirmaRehberi", "S.SEYMEN")`; commit evidence and checkpoint in the same transaction only after a folder completes.

- [ ] **Step 6: Run auth/storage tests**

Run: `python -m pytest tests/test_auth.py tests/test_storage.py -v`
Expected: PASS.

- [ ] **Step 7: Commit authentication and storage**

Run: `git add outlook_rehberi/auth.py outlook_rehberi/storage.py tests/test_auth.py tests/test_storage.py && git commit -m "feat: add secure multi-account sessions and local cache"`

### Task 7: Scan Orchestration, Checkpoint Safety, and Progress

**Files:**
- Create: `outlook_rehberi/scanner.py`
- Create: `tests/test_scanner.py`

**Interfaces:**
- Consumes: Graph, parser, directory, storage, and Task 1 scan models.
- Produces: `ScanCallbacks(on_progress: Callable[[ScanProgress], None], on_warning: Callable[[ScanWarning], None])`
- Produces: `ScanService.run(selections: Sequence[ScanSelection], full_rebuild: bool, cancellation: CancellationToken, callbacks: ScanCallbacks) -> ScanResult`
- Produces: `ScanService.load_folder_trees(accounts: Sequence[AccountRef], cancellation: CancellationToken) -> Mapping[str, tuple[FolderRef, ...]]`

- [ ] **Step 1: Write failing successful-scan tests**

Use fake Graph/auth/storage dependencies. Assert selected descendants are not scanned twice, all pages are parsed, own addresses are excluded, full rebuild clears first, successful checkpoints commit, and progress counters are monotonic.

- [ ] **Step 2: Add multi-account and failure-focus tests**

Implement `test_multi_account_evidence_is_global_and_own_alias_is_excluded` and `test_incomplete_folder_does_not_commit_checkpoint`. Also assert completed folders survive a later failure, the result is marked incomplete, cancellation stops promptly, and no automatic export eligibility is reported.

- [ ] **Step 3: Run scanner tests and verify failure**

Run: `python -m pytest tests/test_scanner.py -v`
Expected: FAIL because the scanner does not exist.

- [ ] **Step 4: Implement scan orchestration**

Keep orchestration synchronous and UI-independent; Task 8 runs it in a Qt worker thread. Deduplicate selected roots before scanning, stage one folder at a time, and emit immutable progress snapshots.

- [ ] **Step 5: Run scanner tests**

Run: `python -m pytest tests/test_scanner.py -v`
Expected: PASS.

- [ ] **Step 6: Commit scanning**

Run: `git add outlook_rehberi/scanner.py tests/test_scanner.py && git commit -m "feat: orchestrate safe multi-account mailbox scans"`

### Task 8: Cross-Platform PySide6 Desktop UI and Save As Flow

**Files:**
- Create: `outlook_rehberi/ui/__init__.py`
- Create: `outlook_rehberi/ui/main_window.py`
- Create: `outlook_rehberi/ui/folder_tree.py`
- Create: `outlook_rehberi/ui/workers.py`
- Create: `outlook_rehberi/ui/dialogs.py`
- Create: `outlook_rehberi/app.py`
- Modify: `outlook_rehberi/__main__.py`
- Create: `tests/ui/test_main_window.py`
- Create: `tests/ui/test_folder_tree.py`
- Reference: `uygulama.py:429-1119`

**Interfaces:**
- Consumes: all service interfaces from Tasks 1-7.
- Produces: `create_application(config: AppConfig) -> QApplication`
- Produces: `MainWindow` actions for add/disconnect account, refresh folders, select folders, start/cancel scan, full rebuild, clear data, export, and legacy import.
- Produces: `ScanWorker` Qt signals `progress`, `warning`, `completed`, `failed`, and `cancelled`.

- [ ] **Step 1: Write failing folder-tree UI tests**

Assert account/folder hierarchy, tri-state propagation, descendant deduplication, multiple-account selection, zero-selection validation, and refresh/error states.

- [ ] **Step 2: Write failing main-window flow tests**

With mocked services, assert connected accounts list, read-only permission copy, disabled scan while busy, progress values, Cancel, incomplete-scan retry state, “Tarama tamamlandı” dialog, native Save As invocation, editable default filename, cancelled Save As, successful atomic export, and export error display.

- [ ] **Step 3: Run UI tests and verify failure**

Run: `QT_QPA_PLATFORM=offscreen python -m pytest tests/ui -v`
Expected: FAIL because the PySide6 UI does not exist.

- [ ] **Step 4: Implement account and folder selection UI**

Use native Qt widgets and platform styles. Keep business logic in services, all network/export work in workers, and only UI state/signals in UI modules.

- [ ] **Step 5: Implement scan completion and Save As behavior**

On complete results, show the completion summary first; when acknowledged, call `QFileDialog.getSaveFileName` with Task 4's default filename and `.xlsx` filter. Keep an explicit “Excel'i Kaydet” action so a cancelled dialog can be reopened.

- [ ] **Step 6: Run UI and full unit tests**

Run: `QT_QPA_PLATFORM=offscreen python -m pytest -v`
Expected: PASS.

- [ ] **Step 7: Commit the desktop UI**

Run: `git add outlook_rehberi/ui outlook_rehberi/app.py outlook_rehberi/__main__.py tests/ui && git commit -m "feat: add Windows and macOS desktop workflow"`

### Task 9: Platform Packaging, Installers, and Release Automation

**Files:**
- Create: `packaging/OutlookFirmaRehberi.windows.spec`
- Create: `packaging/OutlookFirmaRehberi.macos.spec`
- Create: `scripts/build_macos_icon.sh`
- Create: `scripts/smoke_launch.py`
- Modify: `setup.iss`
- Modify: `.github/workflows/build.yml`
- Delete after parity verification: `OutlookFirmaRehberi.spec`
- Modify: `assets/icon.ico`
- Create during macOS build: `assets/icon.icns`
- Create: `tests/test_packaging_config.py`

**Interfaces:**
- Consumes: console entry point and resources from Tasks 1 and 8.
- Produces: Windows x64 installer `.exe`, macOS Apple Silicon `.dmg`, macOS Intel `.dmg`, SHA-256 checksum files, and optional signed/notarized variants.
- Produces: release configuration through `OUTLOOK_REHBERI_CLIENT_ID`/GitHub repository variable without embedding secrets.

- [ ] **Step 1: Write failing packaging-configuration tests**

Assert both spec files target the same entry point, include required MSAL/keyring/PySide6 resources, set platform icons, exclude `pywin32`, and never request/build a console window in release mode.

- [ ] **Step 2: Run packaging tests and verify failure**

Run: `python -m pytest tests/test_packaging_config.py -v`
Expected: FAIL because the new specs do not exist.

- [ ] **Step 3: Implement Windows packaging**

Build with PyInstaller on `windows-latest`, then Inno Setup to create `OutlookFirmaRehberi_v<version>_Windows_x64_Setup.exe`; conditionally sign when certificate secrets exist.

- [ ] **Step 4: Implement macOS packaging**

Build on Intel and Apple Silicon runners available to GitHub Actions, create `.app`, generate the correct icon, use `hdiutil` for `.dmg`, and conditionally codesign/notarize when Apple secrets exist.

- [ ] **Step 5: Add test/build/release gates**

Run all tests on Windows and macOS before packaging; smoke-launch each app in non-interactive mode; publish artifacts only after tests; attach both platform downloads and checksums on version tags.

- [ ] **Step 6: Run local static packaging validation**

Run: `python -m pytest tests/test_packaging_config.py -v && python scripts/smoke_launch.py --source`
Expected: PASS and source smoke launch exits 0 without opening a window.

- [ ] **Step 7: Commit packaging**

Run: `git add packaging scripts setup.iss .github/workflows/build.yml assets tests/test_packaging_config.py OutlookFirmaRehberi.spec && git commit -m "build: package Outlook directory for Windows and macOS"`

### Task 10: Documentation, Integration Verification, and Migration Closure

**Files:**
- Modify: `README.md`
- Create: `docs/MICROSOFT_SETUP.md`
- Create: `docs/PRIVACY.md`
- Create: `tests/test_integration_directory_export.py`
- Modify or retire: `uygulama.py`

**Interfaces:**
- Consumes: the complete application.
- Produces: a supported-user workflow, Entra registration instructions, platform install guidance, permission/privacy statement, and end-to-end fixture proof.

- [ ] **Step 1: Write the failing end-to-end fixture test**

Feed two accounts, nested folders, incoming/sent/deleted messages, duplicate contacts, Polisan/Poliport domains, personal domains, Gümrük signatures/domains, quoted replies, and one malformed message through fake Graph into `ScanService`, then export and assert the complete acceptance criteria.

- [ ] **Step 2: Run the integration test and verify any remaining failure**

Run: `python -m pytest tests/test_integration_directory_export.py -v`
Expected: FAIL only where integration wiring or migration parity remains incomplete.

- [ ] **Step 3: Complete entry-point wiring and legacy closure**

Wire real config/auth/Graph/storage/scanner/export/UI dependencies in `app.py`. Retire COM execution from the default entry point; keep `uygulama.py` only as a documented legacy reference if deletion would impair v4 maintenance history.

- [ ] **Step 4: Update user and administrator documentation**

Document account support, one-time sign-in, exact read-only permissions, multi-account/folder selection, all-history behavior, workbook categories, Save As flow, v4 migration, consent/admin errors, Windows/macOS installation, unsigned-build warnings, and signing variables.

- [ ] **Step 5: Run all verification**

Run: `QT_QPA_PLATFORM=offscreen python -m pytest -v`
Expected: all tests PASS.

Run: `python -m compileall -q outlook_rehberi tests`
Expected: exit 0.

Run: `git diff --check`
Expected: exit 0.

- [ ] **Step 6: Commit documentation and integration**

Run: `git add README.md docs outlook_rehberi uygulama.py tests/test_integration_directory_export.py && git commit -m "docs: complete cross-platform v5 release guidance"`

- [ ] **Step 7: Perform final branch review**

Compare every goal and acceptance statement in the spec to code, tests, and release artifacts. Record the production Entra client ID and signing credentials as external release prerequisites rather than claiming signed downloads exist without them.

