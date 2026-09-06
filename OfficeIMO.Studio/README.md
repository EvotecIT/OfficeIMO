# OfficeIMO Studio

OfficeIMO Studio is the cross-platform desktop surface for OfficeIMO's document engines. The application presents the existing PDF, conversion, workflow, security, and OCR capabilities; reusable document behavior remains in the owning OfficeIMO packages.

## Current workspace

- Open several PDFs in document tabs and use single-page, continuous, two-page, or grid reading modes.
- Opening a file through a symlink or hard link reuses its existing tab and pending edits. Conversion, document-health output, assembly, page export, Save As, extraction, splitting, and protected/decrypted copies check live tab ownership before publishing their final destination. Numbered-copy mode skips owned paths; replacement reports a failure instead of replacing an open document. Assembly adds each physical source once, and comparison rejects an alias of the current document.
- Local-file saves check that the source still has the physical identity and contents Studio opened. If another application changes, replaces, or removes it, Studio keeps the pending edits and refuses to overwrite the external version; use **Save As** to choose a different file. The shared commit path also checks the displaced source during atomic replacement. Saving a source opened through a symlink updates its target and preserves the link.
- Open and compare PDFs through desktop storage-provider streams, including sources without local paths. Available provider bookmarks are retained with recent and restart references. PDF input and serialized provider output are limited to 512 MiB. A provider save checks the opened contents before writing and verifies the result after closing the write stream, but cannot guarantee atomic replacement or rollback. Studio explains this before saving and keeps edits and enabled recovery snapshots when publication fails.
- Choose comfortable or compact controls in Settings. With document history enabled, Studio remembers page, zoom, reading layout, bounded pane widths, and pane choices for up to 64 documents, using path hashes in the view-preference file.
- Press F9 for focus reading and Escape or F9 to restore document controls. At narrow widths, a document-task picker and one optional side pane leave room for the page.
- Custom zoom stays in place through editing, undo, redo, recovery, and saving. Fit width and Fit page remain explicit choices.
- Continue a previous session from Home. Unchanged sources reopen with their reading positions; changed or missing sources offer individual choices, including recovering verified pending edits into a new file. Recovering to a local file requires a new path; provider destinations use the explicit direct-write confirmation. Session memory can be disabled in Settings and keeps up to 32 document references locally; entries expire after 30 days of inactivity.
- If Studio cannot save or clear the restart record, a notice stays visible with a Retry action. A successful retry clears that notice without dismissing unrelated recovery errors.
- Turn off document history in Settings to stop storing new recent-file entries and reading preferences. Existing records remain until cleared. Before working with private documents, also disable restart sessions and recovery storage, and close other running Studio instances. Confirmed history cleanup removes recent files, saved reading preferences, and restart records; it preserves open documents, unsaved edits, and recovery copies. Failed or partial cleanup is reported for retry.
- Recovery restores only snapshots written within the last 30 days and verifies their contents again when applying them. Each new snapshot stores its metadata and document bytes in one atomically replaced file, preserving the previous snapshot if publication fails. Existing recovery files remain readable and migrate on the next successful edit while storage is enabled. Recovery snapshots are limited to 512 MiB; with recovery storage enabled, an edit that would exceed this limit is rejected before changing the open document.
- Turn off recovery storage in Settings before editing private documents to prevent new recovery copies. This waits for active writes and applies to the running Studio instance and future launches; close other running instances separately. Editing, undo, redo, and saving remain available, but new unsaved edits cannot be recovered after an unexpected exit. Existing copies remain available until explicitly cleared or expired. This switch does not disable recent documents, saved reading positions, or restart sessions.
- Studio cleans recognized expired recovery data at startup and removes abandoned temporary files older than one day. Automatic cleanup retains unknown future formats and unrelated files. Settings also offers confirmed cleanup of stored recovery files and reports any files it cannot remove. Cleanup preserves original documents and open unsaved edits; subsequent edits can create new recovery files while storage is enabled.
- On Unix, newly written recovery files allow read/write access only to their owner. Rewriting an older snapshot also applies these permissions. Windows recovery files use the access controls of the local application-data location.
- Search the Tools catalog or open command search with Ctrl+Shift+P (Command+Shift+P on macOS). Commands show document requirements, use the same operations as the workspace controls, and can be selected with the keyboard.
- Search, follow bookmarks and links, navigate by keyboard, use page night mode, and compare two PDFs in synchronized panes.
- Select existing text, images, and annotations with visible bounds and handles. Replace, move, resize, recolor, flatten, or remove supported objects through the canonical PDF editors.
- Add text, images, links, notes, markup, shapes, ink, stamps, signature appearances, watermarks, and page numbers.
- Insert or replace PNG and JPEG images through the file picker's stream access, including sources without a local path. Each selected image is limited to 64 MiB; access is released after reading, and failed reads preserve the current document.
- Organize pages by drag and drop; rotate, crop, duplicate, reorder, import, extract, split, and insert blank pages with undo, redo, save, and recovery support. PDF imports accept provider streams and apply the selected batch as one undoable change, with a 512 MiB total input limit.
- Fill, flatten, and author supported AcroForm controls.
- Inspect protection and signatures; protect or decrypt a copy; apply certificate signatures, validate signatures, add Bates numbering, sanitize, repair, optimize, and perform verified redaction.
- Convert supported document formats, export PDF pages to supported image formats, preview print sheets, and assemble PDFs from ordered PDFs, images, Office files, folders, and ZIP archives. Conversion and assembly accept provider-backed files and ZIPs through bounded private staging, retain their original references, and check provider contents again before publication. Choose a filesystem destination for these workflows. OCR, print preview, page export, and folder intake currently require filesystem locations.
- In the conversion queue, **Run pending** processes new entries and **Retry failed / cancelled** retries only those outcomes, using the current output settings. Completed entries keep their files and diagnostics. Aliases of a file are queued once per conversion route, while distinct files on case-sensitive volumes remain separate. If a runner stops without returning a result, **Check output** requires inspecting the destination before removing and re-adding that entry; it is excluded from automatic retries. To intentionally repeat a completed conversion, remove its queue entry and add the source again.
- Open **Jobs** to follow conversion, document-health, assembly, page-export, and searchable-PDF attempts across document tabs. These workflows share two execution slots; conversion files run sequentially within their batch. Cancel a waiting or running attempt from its card, or cancel its conversion batch. Completed PDF outputs open in Studio; other files and folders use the system association. The history holds up to 500 attempts in memory for the current session, retaining active entries and discarding the oldest finished entries when full. **Clear finished** removes records without deleting output files. This history does not restore or resume execution after restarting Studio.
- Create searchable PDFs with the public `OfficeIMO.Reader.Ocr` facade. The workbench supports page ranges, 28 typed language choices, confidence and rendering controls, checksum-verified language-data provisioning, cancellation, and safe output conflict policy.
- Closing a tab or Studio while work is active offers **Wait and close**, **Cancel work and close**, and **Keep open**. Waiting or cancelling keeps the affected document alive until operations finish; cancelling one tab leaves other tabs' jobs running. Whole-window close also waits for session restoration. Cancelling restoration retains unopened entries for a later restart choice. Already saved outputs are preserved, and unsaved-edit choices follow once active work has stopped. An unexpected process exit does not resume workflow execution.

## Capability ownership

Studio does not contain a second document engine:

- `OfficeIMO.Pdf` owns PDF reading, rendering, interaction geometry, editing, forms, security, signatures, redaction, and output mutation.
- `OfficeIMO.Workflows` owns reusable conversion, inspection, repair, optimization, comparison, sanitization, image-export, print-planning, and mixed-source assembly workflows.
- `OfficeIMO.Reader.Ocr` owns the easy searchable-PDF flow over the engine-neutral Reader OCR contracts and the optional Tesseract CLI provider.
- Avalonia owns only the cross-platform windowing and control layer. Studio renders OfficeIMO's retained page scene; it does not use PDFium or convert pages to images merely to display them.

Tesseract is an explicit OCR runtime prerequisite. OfficeIMO discovers an installed executable and can provision checksum-pinned language data, but the application does not silently install or bundle a native OCR executable.

## Language, accessibility, and local data

Studio stores versioned user preferences under the operating system's local application-data directory. English is the fallback interface language. The language catalog already reserves Polish, German, French, Italian, and more than fifteen additional culture packs; only reviewed translations appear in the settings picker. The `en-XA` expansion locale is available now to expose clipped controls, hard-coded strings, and layouts that cannot accommodate longer translations.

User-facing XAML, dialogs, file pickers, and capability notices resolve stable, feature-scoped resource keys through `IStudioLocalizer`. Document engines remain culture-neutral. A new translation adds satellite resources and a reviewed catalog entry rather than branching view models or document operations by language.

Studio follows Avalonia's platform font selection and fallback instead of forcing a Windows-only typeface. System, light, dark, and explicit high-contrast appearance preferences share the same semantic color resources. Navigation and editing controls expose automation names, icon-only controls include help text, normal tab order follows the visual flow, and the shell supports the primary platform modifier for Open, Find, tab switching, tab closing, zoom, and fit commands. Page navigation also supports arrows, Page Up/Down, Home, and End.

Crash diagnostics stay local, bounded, and privacy-safe. They record stable event codes, runtime metadata, exception type/HResult, and sanitized stack frames. They do not record document contents, document names, document paths, or exception messages.

## Distribution

PowerForge owns the release matrix, signing, archives, generated Windows MSI, Debian package, macOS `.app`, checksums, and artifact manifests. Validate or inspect the repository-local product configuration with:

```powershell
./Build/Studio/Build-Studio.ps1 -Validate
./Build/Studio/Build-Studio.ps1 -Plan
./Build/Studio/Build-Studio.ps1 -Target Studio.Windows -Runtime win-x64
```

The initial update policy is manual: install a newer signed artifact over the stable application identity. Building artifacts never publishes them. See [`Build/Studio/README.md`](../Build/Studio/README.md) for runtime targets and the current native-package boundaries.

## Host boundary

Studio's reusable document behavior stays in OfficeIMO packages, while presentation preferences, localization, accessibility, and transport-neutral activation belong to the Studio host layer. A future browser companion should send a bounded activation request through an authenticated local native host and let Studio open the document or workflow. It should not duplicate PDF editing, OCR, conversion, storage, or policy logic in an extension.

## Run and verify

```powershell
dotnet run --project OfficeIMO.Studio/OfficeIMO.Studio.csproj
dotnet test OfficeIMO.Studio.Tests/OfficeIMO.Studio.Tests.csproj -c Release
```

The document workspace currently opens PDF files directly. Other supported formats enter through conversion or mixed-source assembly. Print planning and preview are implemented; native operating-system printer submission remains open work.

## Next product outcomes

The repository's single open-work plan is [`Docs/ROADMAP.md`](../Docs/ROADMAP.md). Its **Desktop Studio** section tracks distribution, richer OCR review, native print and scan intake, reusable workflow recipes, and cross-platform accessibility and usability evidence.
