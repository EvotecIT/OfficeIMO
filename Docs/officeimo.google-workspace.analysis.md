# OfficeIMO Google Workspace: Current State and Roadmap

Status: working roadmap

Snapshot: 2026-07-10

Scope: Google Drive, Docs, Sheets, and Slides integrations owned by OfficeIMO

## Decision

OfficeIMO should continue the Google work as a family of extension libraries, but the next investment should harden the packages that already exist before adding more mapping breadth.

The target package family is:

- `OfficeIMO.GoogleWorkspace`: dependency-light session, token, transport, error, retry, scope, and fidelity contracts.
- `OfficeIMO.GoogleWorkspace.Drive`: a typed Drive owner used by every document translator.
- `OfficeIMO.GoogleWorkspace.Auth.GoogleApis`: an optional adapter for Google client credentials and interactive OAuth without forcing that dependency into the core package.
- `OfficeIMO.Word.GoogleDocs`: Word and Google Docs translation.
- `OfficeIMO.Excel.GoogleSheets`: Excel and Google Sheets translation.
- `OfficeIMO.PowerPoint.GoogleSlides`: PowerPoint and Google Slides translation.

The existing Word and Excel translators already use domain-specific inspection snapshots and neutral request plans. Keep that design. Do not create a universal `OfficeIMO.Documents` intermediate representation until two implemented translators prove a concrete shared semantic model. Word, spreadsheets, and presentations are too different to justify that abstraction today.

The first `1.0` releases should mean that a package can analyze, create, read/import, and safely replace its Google-native document with honest fidelity reporting and live end-to-end tests. Fine-grained two-way synchronization should follow in `1.x`; it depends on stable importers and identity markers and should not hold the first production contract hostage.

## What Exists Today

All three current packages are published on NuGet at `0.1.15` as of this snapshot.

| Package | Current owner | Current maturity |
| --- | --- | --- |
| `OfficeIMO.GoogleWorkspace` | Credentials, session options, scopes, retry, diagnostics, and Drive location/reference primitives | Useful shared kernel, but not a complete Google or Drive client |
| `OfficeIMO.Word.GoogleDocs` | Word inspection, translation planning, Google Docs request compilation, create/replace export, and image/header/footer execution | Substantial one-way exporter with important contract, security, modern-Docs, and import gaps |
| `OfficeIMO.Excel.GoogleSheets` | Workbook inspection, translation planning, Sheets request compilation, and create/replace export | Useful one-way tabular exporter; advanced spreadsheet and import paths remain incomplete |
| `OfficeIMO.GoogleWorkspace.Tests` | Credential, failure, retry, and transport-boundary contracts | Mocked HTTP evidence only; no disposable live Google validation lane |
| `OfficeIMO.PowerPoint.GoogleSlides` | Not present | Missing product-family surface |

The current execution shape is sound:

```text
OfficeIMO.Word       -> inspection snapshot -> Docs translation plan -> neutral Docs batch -> Docs/Drive REST
OfficeIMO.Excel      -> inspection snapshot -> Sheets translation plan -> neutral Sheets batch -> Sheets/Drive REST
OfficeIMO.PowerPoint -> missing
```

The plans are useful seams for dry-run diagnostics, payload tests, and future diff planning. The problem is not the overall direction. The problem is that parts of the public contract run ahead of the actual execution path.

## Current Capability Inventory

### Shared Google Workspace kernel

Implemented:

- Static access-token credentials.
- Application-supplied asynchronous credential delegates.
- Service-account JWT token exchange, token caching, and domain-wide delegation subject support.
- Docs, Sheets, and Drive OAuth scope constants.
- Session-level `HttpClient`, timeout, retry, default folder, and diagnostics configuration.
- Retry handling for 408, 429, and common 5xx responses, including `Retry-After` and jittered exponential backoff.
- Typed cancellation, timeout, token-acquisition, delegation, and API failure outcomes with a translation report.
- `GoogleDriveFileLocation` and common file-reference results.

Not implemented or incomplete:

- Interactive installed-app or web OAuth, PKCE, refresh-token persistence, revocation handling, and incremental consent.
- A reusable typed HTTP transport; Docs and Sheets contain parallel send/error/Drive-placement implementations.
- A typed Drive client for files, folders, shared drives, permissions, comments, revisions, exports, uploads, and change tracking.
- Least-privilege scope selection per operation beyond the current authoring bundles.
- Request identity and product headers: `ApplicationName` is public configuration but is not used in outgoing requests.
- Real shared-drive targeting: `DriveId` influences a diagnostic/query flag, but a concrete folder is still required and the drive is not resolved or validated.
- Typed Google error reasons, quota classification, request IDs, resumable uploads, or safe retry classification per operation.
- A stable diagnostic code/catalog, source object paths, remediation metadata, or strict fidelity policies.

### Word to Google Docs

Implemented end-to-end through the current exporter:

- Offline translation plan and compiled neutral batch.
- Create a Google Doc or destructively replay into an existing document ID.
- Move a created document into a configured Drive folder.
- Paragraphs and headings.
- Bold, italic, underline, strike, font size/family, foreground/highlight colors, baseline offsets, and small caps.
- External hyperlinks.
- Paragraph alignment, indentation, before/after spacing, line-spacing approximations, right-to-left direction, pagination controls, shading, borders, and tab stops.
- Ordered and bulleted lists.
- Page breaks and native section breaks.
- Page size, margins, header/footer distance, columns, first-page header/footer behavior, and page-number start where the Docs model permits it.
- Default, first-page, and even-page headers and footers, including simple tables.
- Body and table footnotes.
- Bookmarks emitted as named ranges across body, headers/footers, footnotes, and table cells.
- Tables, merged cells, repeated header rows, column widths, cell shading, and supported borders.
- Inline images in body, table, header/footer, and footnote content.
- Retry and failure diagnostics around Docs and Drive calls.

Partially implemented or intentionally lossy:

- All-caps run formatting is detected but not written; small caps is written.
- Exact and at-least line spacing use an approximation; other rules remain unsupported.
- Tab-stop leaders have no current target mapping.
- Even/odd section breaks fall back to `NEXT_PAGE`.
- Internal links retain source anchor data, but the final tab-aware bookmark/heading link is not emitted.
- Existing-document replacement clears and replays content instead of reconciling it. It has no collaborator/revision guard.
- Header/footer, table, and footnote replay depends on live index discovery over several request rounds.

Detected but not actually flattened, rasterized, or exported:

- Floating shapes and text boxes.
- Word charts.
- SmartArt.
- Content controls.
- Embedded OLE objects.
- Watermarks.
- Equations.
- Comments.
- Nested-table fidelity beyond diagnostics.

Missing platform coverage:

- Google Docs tabs and child tabs. Current reads used during export operate on the legacy first-tab shape and do not request all tab content.
- Tab-aware bookmark and heading links.
- `WriteControl` revision protection for collaborative documents.
- Google Docs to `OfficeIMO.Word` import.
- Drive-export-to-DOCX convenience import.
- Native read/update/diff APIs.
- Suggestions/revisions policy.
- Permission, sharing, revision, export, and change-feed operations.

### Excel to Google Sheets

Implemented end-to-end through the current exporter:

- Offline translation plan and compiled neutral batch.
- Create a spreadsheet or destructively rebuild an existing spreadsheet ID.
- Move a created spreadsheet into a configured Drive folder.
- Sheet order, names, visibility, right-to-left display, tab colors, and frozen rows/columns.
- Row heights, column widths, and hidden rows/columns.
- Blank, string, numeric, Boolean, date/time, and formula cell values.
- Basic number formats, bold, italic, underline, font/fill color, borders, alignment, and wrapping.
- External hyperlinks and resolvable internal sheet/named-range links.
- Cell comments flattened to cell notes.
- Merged ranges.
- User-defined named ranges; built-in names such as print areas remain diagnostics.
- Whole-sheet protection, without full Excel per-operation permission parity.
- Basic filters, additional filter views, value filters, and a useful subset of custom text/numeric filters.
- Native Google Sheets tables, header/banding/footer colors, totals-row cells, and table column metadata.
- List, numeric, date, and text-length validation for populated and validation-only cells.
- Retry and failure diagnostics around Sheets and Drive calls.

Partially implemented or intentionally lossy:

- Formulas are normalized by adding `=` and sent as-is. There is no Excel-to-Sheets function compatibility catalog or fallback execution.
- Table style names and totals-function metadata are mainly diagnostic.
- Worksheet protection is coarse.
- Complex custom filter combinations fall back to diagnostics.
- Rich Excel comments are plain notes in Sheets.
- Formatting omits material cell properties such as font size/family, strike, rotation, indentation, rich text runs, themes, and several number-format edge cases.

Options or plan concepts without an execution path:

- Charts.
- Pivot tables.
- Header/footer metadata.
- Fallback of unsupported formulas to text.

Missing spreadsheet coverage:

- Conditional formatting. The old analysis listed it as a first delivery feature, but no request model or payload exists today.
- Charts, pivot tables, slicers, banded ranges, row/column groups, developer metadata, and data-source objects.
- Smart chips and rich cell text runs.
- Images and drawing objects.
- Locale, time zone, calculation settings, themes, hidden gridlines, and broader sheet properties.
- High-throughput `spreadsheets.values.batchUpdate`, range chunking, sparse-write planning, and large-workbook quota strategy.
- Google Sheets to `OfficeIMO.Excel` import.
- Drive-export-to-XLSX convenience import.
- Native read/update/diff APIs.
- Permission, sharing, revision, export, and change-feed operations.

### Google Slides

There is no OfficeIMO Google Slides package, even though OfficeIMO has a first-party PowerPoint model. This is now the largest product-family omission.

The Slides API can create and modify presentations, slides, shapes, text, tables, images, video, charts, transforms, ordering, and speaker-note text. That makes a direct `OfficeIMO.PowerPoint` translator viable. It should follow the proven plan/batch pattern rather than waiting for a cross-document intermediate model.

## Contract and Operational Risks to Fix First

### 1. Public options imply behavior that does not happen

| Surface | Current behavior |
| --- | --- |
| `GoogleDocsSaveOptions.FlattenFloatingContent` | Changes a diagnostic sentence; it does not flatten content |
| `GoogleDocsSaveOptions.RasterizeWordCharts` | Changes a diagnostic sentence; it does not rasterize charts |
| `GoogleDocsSaveOptions.PreserveCommentsViaDriveApi` | Changes a diagnostic sentence; it does not create comments |
| `IncludeHeadersAndFooters`, `IncludeFootnotes`, `IncludeBookmarksAsNamedRanges` | Not read; those features are compiled regardless of the option |
| `GoogleSheetsSaveOptions.IncludeCharts`, `IncludePivotTables` | Change planning diagnostics only; no chart or pivot requests exist |
| `IncludeHeaderFooterMetadata` | Not read |
| `PreserveUnsupportedFormulasAsText` | Changes a diagnostic sentence; formulas are still sent as formulas |
| `TreatPrintLayoutAsDiagnosticOnly` | Changes wording only |
| `GoogleWorkspaceSessionOptions.ApplicationName` | Not sent or otherwise used |

These are preview APIs, so intentional cleanup is preferable to preserving misleading compatibility. Every option must either alter execution, select a documented fidelity policy, or be removed before `1.0`.

### 2. Inline image staging leaves public Drive files behind

The Docs exporter uploads each image to Drive, creates an `anyone/reader` permission, inserts the public URL, and does not revoke the permission or delete the staging file. This is a confirmed code-path risk, not a theoretical roadmap item.

The safe model is an image content lease:

1. Obtain a URI the Docs API can fetch.
2. Record every temporary file and permission created by the operation.
3. Insert the images.
4. Revoke temporary public access and delete or retain the staging object according to an explicit policy, even after partial failure.

Until that exists, public-link staging should require explicit opt-in and produce a high-severity diagnostic.

### 3. Replace operations are destructive and collaboration-blind

Both exporters rebuild existing content. Docs replacement does not use `WriteControl`, and Sheets replacement does not expose a conflict policy. A concurrent editor can lose work or leave the exporter working against stale indexes.

Separate these operations in the public API:

- `Create`: always creates a new Google-native file.
- `Replace`: explicit destructive replacement, guarded by the best available revision/version precondition.
- `Update` or `Synchronize`: plans a diff, reports conflicts, and applies only accepted changes.

### 4. Retry policy does not distinguish safe and unsafe requests

The shared policy retries POST creates, multipart uploads, and batch updates after server errors, network failures, and timeouts. Google recommends exponential backoff, but also warns that retrying non-idempotent operations is not always safe. A response-lost create can produce duplicates.

The transport needs request semantics:

- always safe: reads and explicitly idempotent updates;
- conditionally safe: operations with a precondition or stable request identity;
- unsafe by default: creates, uploads, permission grants, and destructive batch mutations.

Unsafe retries need operation-specific recovery or duplicate detection, not a global status-code decision.

### 5. Tests prove payloads, not Google behavior

The current suite has strong local coverage: 15 shared credential/transport facts, 78 Word/Docs facts, and 25 Excel/Sheets facts in the principal Google test files. Most API responses are hand-authored mock JSON. That proves serialization and orchestration but does not prove Google accepts the payloads or renders the intended result.

The roadmap needs a disposable live test lane with cleanup, plus Drive export/readback comparison. Secrets stay in CI or the operator environment; recorded responses must be scrubbed.

### 6. Package and website messaging is inconsistent

The packages contain real exporters, while project and website descriptions still call them “scaffolding.” The old analysis still recommends creating packages that have been shipping since March 2026. There are package READMEs and examples, but no coherent Google Workspace section on the product website and no maintained support matrices.

## Target Architecture

### `OfficeIMO.GoogleWorkspace`

Keep the kernel small and reusable:

- `IGoogleWorkspaceCredentialSource` and access-token results.
- Session and per-request context.
- Scope catalog and least-privilege scope composition.
- Typed HTTP transport with product headers, cancellation, timeout, quota user, request IDs, JSON, media, and resumable upload support.
- Typed Google error payloads and retry classifications.
- Fidelity report, stable diagnostic codes, source paths, target actions, and strictness policy.
- Clock/delay abstractions where deterministic tests need them.

Do not put Docs, Sheets, Slides, or full Drive domain models in this package.

### `OfficeIMO.GoogleWorkspace.Auth.GoogleApis`

Make Google client-library integration optional:

- adapt `GoogleCredential`/token access to `IGoogleWorkspaceCredentialSource`;
- support desktop OAuth with system browser and PKCE;
- provide token-store contracts and secure-store examples, not a plaintext default hidden in the library;
- handle refresh, invalidation, revoked consent, and incremental scopes;
- retain the dependency-light service-account source in the core package unless the optional adapter clearly replaces it.

Applications still own client IDs, client secrets, consent-screen configuration, and tenant policy.

### `OfficeIMO.GoogleWorkspace.Drive`

This becomes the single owner for cross-product Drive work:

- file/folder metadata, capabilities, list/get/search, create/copy/move/delete, and shared-drive resolution;
- permissions and sharing;
- comments and replies, with the Google editor anchoring limitation documented;
- revisions and exports;
- simple, multipart, and resumable uploads;
- Office-to-Google import conversion and Google-to-Office export;
- change tokens for users and each relevant shared drive;
- temporary content leases for images and other externally fetched media;
- cleanup outcomes that survive partial failures.

Docs, Sheets, and Slides should call this client instead of owning Drive URLs and duplicate metadata payloads.

### Domain translators

Each translator owns only its domain mapping:

```text
OfficeIMO model <-> domain inspection snapshot <-> domain translation plan <-> Google API model
```

Common operation families should be recognizable without forcing identical models:

- analyze/build plan without network access;
- create;
- explicit destructive replace;
- import through native Google resources;
- import/export through Drive format conversion as a broad fallback;
- plan and apply synchronization after stable identity mapping exists.

Every result should return the Google file ID, URL, fidelity report, remote revision/version evidence when available, created temporary resources, and cleanup status.

## Fidelity Contract

Replace feature-specific booleans that do not execute with a small policy model:

```csharp
public enum UnsupportedFeatureMode {
    Error,
    WarnAndSkip,
    Flatten,
    Rasterize
}
```

Not every translator must support every mode. A feature capability table should state which modes are legal. For example, equations might support `Error`, `WarnAndSkip`, and later `Rasterize`, while comments might support `Error`, `WarnAndSkip`, or Drive comment creation.

Each diagnostic needs:

- stable code, such as `DOCS.CHART.RASTERIZED`;
- source path or stable source identifier;
- severity;
- selected action;
- count and optional target identifier;
- concise remediation or option guidance.

Add a strict preflight mode that fails before any Google mutation when the plan contains an unaccepted lossy mapping.

## Phased Roadmap

### Phase 0 — Make the preview contract honest and safe

- [ ] Replace public image staging with a temporary-content lease, cleanup, and explicit opt-in policy.
- [ ] Wire or remove every no-op/misleading option listed above.
- [ ] Rename “scaffolding” package descriptions to match implemented behavior while clearly marking preview limitations.
- [ ] Add stable diagnostic codes, source paths, selected actions, and a strict preflight mode.
- [ ] Move duplicated HTTP/error/Drive-placement code into the shared transport.
- [ ] Classify retries by operation safety; add recovery/deduplication tests for ambiguous creates and uploads.
- [ ] Send `ApplicationName` and other supported request context, or remove the unused setting.
- [ ] Publish current Docs and Sheets support matrices from code-owned source data.
- [ ] Add API compatibility baselines and package validation for the three published packages.
- [ ] Split the 4,300-line Word/Docs test file into capability-focused files without weakening contracts.

Exit criteria:

- no public option lies about execution;
- no default export leaves an `anyone`-readable staging file;
- retry behavior is safe for mutation type;
- package copy and support matrices describe current behavior;
- payload and public API compatibility are gated in CI.

### Phase 1 — Build the shared Drive and authentication foundation

- [ ] Add `OfficeIMO.GoogleWorkspace.Drive` with file metadata, capabilities, folder/shared-drive resolution, copy/move/delete, and permissions.
- [ ] Add import/export format discovery through Drive `about.importFormats` and `about.exportFormats`.
- [ ] Add small-file and resumable upload/download/export APIs with progress and cancellation.
- [ ] Add comments/replies, revisions, and change-token clients.
- [ ] Add temporary media leases and cleanup reports.
- [ ] Add the optional Google client credential adapter with desktop PKCE and refresh-token support.
- [ ] Let callers choose minimum scopes per operation and surface consent requirements before mutation.
- [ ] Replace Docs/Sheets Drive code with the new owner.
- [ ] Create a disposable live-test fixture that creates, reads, exports, and deletes files in a dedicated folder.
- [ ] Add a conservative shared-drive test lane when credentials and a test drive are configured.

Exit criteria:

- Docs and Sheets contain no private duplicate Drive client;
- user OAuth and service-account flows both have documented production paths;
- live tests clean their files even after failure;
- shared-drive targeting resolves and verifies the actual drive/folder rather than toggling a query flag.

### Phase 2 — Take Google Sheets to a production `1.0`

#### Complete the current exporter

- [ ] Add an Excel-to-Sheets formula compatibility catalog, function/reference rewriting, unsupported detection, and a real fallback policy.
- [ ] Add missing cell formatting: font family/size/strike, rotation, indentation, rich text runs, theme/color handling, and verified number formats.
- [ ] Add conditional formatting with an explicit supported-rule matrix.
- [ ] Add spreadsheet and sheet settings: locale, time zone, calculation properties, gridlines, themes, and remaining visibility/display properties.
- [ ] Expand protection mapping to editors, warning-only ranges, unprotected subranges, and capability-aware diagnostics.
- [ ] Define comment behavior explicitly: note flattening in `1.0`; Drive/Workspace comment semantics only where the platform supports them.
- [ ] Use `spreadsheets.values.batchUpdate` for value-heavy writes and structural `batchUpdate` for formats/objects.
- [ ] Add sparse-range planning, request chunking, payload limits, and quota-aware progress for large workbooks.

#### Add advanced spreadsheet objects

- [ ] Translate supported chart families and publish a chart-property matrix.
- [ ] Translate supported pivot tables and publish a pivot-feature matrix.
- [ ] Add row/column groups, banded ranges, slicers, and developer metadata.
- [ ] Add smart chips where the source has a real Drive/person link semantic; do not infer chips from arbitrary text.
- [ ] Decide image behavior only after a supported API path is proven. `IMAGE()` formulas and Drive links are not equivalent to embedded Excel drawings.

#### Add read/import and safe update

- [ ] Add a broad Drive-export-to-XLSX import path as the first complete fallback.
- [ ] Add native `spreadsheets.get`/values import for sheets, values, formulas, styles, named ranges, filters, tables, validation, charts, and pivots in supported order.
- [ ] Add partial/range import and field masks for large spreadsheets.
- [ ] Add stable OfficeIMO identity markers through developer metadata where appropriate.
- [ ] Add a diff plan that separates source changes, remote changes, conflicts, and lossy actions before apply.
- [ ] Add live round-trip checks: OfficeIMO -> Sheets -> XLSX -> OfficeIMO, plus native API readback.

Sheets `1.0` gate:

- create, explicit replace, native read/import, and Drive-export import are stable;
- formula and advanced-object support is discoverable before mutation;
- large-workbook behavior is chunked and observable;
- live tests cover values, formulas, styles, merges, names, filters, tables, validation, conditional formatting, charts, and pivots at their documented support level.

### Phase 3 — Take Google Docs to a production `1.0`

#### Catch up with the current Docs model

- [ ] Add tab and child-tab resource models and always choose an explicit tab strategy.
- [ ] Make bookmarks, headings, and internal links tab-aware using the current link model.
- [ ] Add `WriteControl` with required/target revision policies for replace and update operations.
- [ ] Make existing-document replacement account for tabs, headers, footers, footnotes, named ranges, and stale segments instead of clearing only the main body path.
- [ ] Reduce staged read/write rounds where object IDs or ordered batches can avoid them.

#### Finish fidelity behavior already advertised by options

- [ ] Implement or remove all-caps, tab-leader, and remaining line-spacing promises.
- [ ] Implement comments through the shared Drive client with the documented unanchored-editor limitation.
- [ ] Add real flatten/rasterize adapters for charts, shapes, text boxes, SmartArt, equations, watermarks, and content controls where OfficeIMO already owns a renderer.
- [ ] Keep unsupported OLE and layout cases explicit; never silently drop them.
- [ ] Add a safe image URI lease and verify permission cleanup in live tests.

#### Add read/import and safe update

- [ ] Add a broad Drive-export-to-DOCX import path as the first complete fallback.
- [ ] Add native Docs import for tabs, body content, styles, lists, tables, images, headers/footers, footnotes, named ranges, and supported links.
- [ ] Add configurable handling for suggestions and revisions.
- [ ] Preserve stable source identity through named ranges/bookmarks where that is semantically safe.
- [ ] Add a diff plan and conflict result before any incremental synchronization apply.
- [ ] Add live round-trip checks: OfficeIMO -> Docs -> DOCX -> OfficeIMO, plus tab-aware native API readback.

Docs `1.0` gate:

- create, explicit revision-guarded replace, native read/import, and Drive-export import are stable;
- all tabs are handled or rejected by explicit policy;
- every lossy Word feature has an executed fallback or preflight diagnostic;
- no image staging object remains public after successful or failed export;
- live tests cover the supported body, table, image, header/footer, footnote, section, bookmark, and tab scenarios.

### Phase 4 — Add `OfficeIMO.PowerPoint.GoogleSlides`

Start only after the shared transport, Drive client, fidelity contract, and live-test fixture exist.

#### First preview

- [ ] Add `GoogleSlidesTranslationPlan`, a neutral Slides batch, and `IGoogleSlidesExporter`.
- [ ] Create presentations and explicitly replace existing content.
- [ ] Map slide order, size, backgrounds, layouts/placeholders, text boxes, paragraphs/runs, hyperlinks, basic shapes, tables, and images.
- [ ] Map speaker-note text, with the API's notes-page limits documented.
- [ ] Support deterministic object IDs inside a batch without treating them as permanent synchronization identities.
- [ ] Add chart mapping for supported PowerPoint chart families and image fallback for the rest.
- [ ] Add video links only where the source and target semantics match.

#### Import and production contract

- [ ] Add Drive-export-to-PPTX import.
- [ ] Add native Slides import for supported pages and page elements.
- [ ] Add template-based generation from an existing presentation.
- [ ] Add revision/conflict policy and diff planning.
- [ ] Publish explicit limits for masters/themes, transitions, animations, diagrams, equations, embedded objects, and unsupported media.
- [ ] Add live create/read/export/delete and round-trip tests.

Slides should remain preview until create, import, safe replace, fidelity reporting, and live validation meet the same package contract as Docs and Sheets.

### Phase 5 — Synchronization and ecosystem polish

- [ ] Add change-token consumption for the user and every relevant shared drive.
- [ ] Store only the minimum synchronization checkpoint and stable identity metadata; do not build a second document database in these packages.
- [ ] Expose plan/apply workflows with conflict, dry-run, cancellation, and partial-failure outcomes.
- [ ] Add website pages for Google Workspace, Docs, Sheets, Slides, Drive, authentication, shared drives, and live-test setup.
- [ ] Publish generated support matrices and runnable installed-package examples.
- [ ] Add migration notes for preview API cleanup and each `1.0` package.
- [ ] Track official API release notes and run schema/payload drift checks regularly.

## Recommended First Implementation Train

The first implementation sequence should stay narrow enough to review and release:

1. **Contract and image safety:** remove/wire no-op options, make public staging opt-in, add cleanup tracking, and correct package copy.
2. **Shared transport extraction:** move request/error/retry behavior out of Docs and Sheets and classify mutation safety.
3. **Drive owner and live harness:** add folder/file/capability/export primitives and disposable create/read/export/delete tests.
4. **Sheets completeness slice:** formula policy, conditional formatting, missing core styles, and values batching.
5. **Docs modern-model slice:** tabs, tab-aware links, revision controls, and complete replace semantics.
6. **Import slice:** Drive-format fallback import first, then native importers.
7. **Slides preview:** reuse the proven shared foundation and domain-plan pattern.

This order fixes current-user risk before expanding the product surface, gives Sheets and Docs credible `1.0` gates, and prevents the Slides package from copying the transport and Drive problems already visible in the first two translators.

## Explicit Non-Goals

- Pixel-perfect parity between Microsoft Office and Google editors.
- Silent preservation of features for which Google exposes no equivalent.
- A universal Office document intermediate representation before real shared semantics emerge.
- Bundling application OAuth secrets or choosing tenant consent policy for callers.
- Treating Drive conversion as proof of native API fidelity.
- Promising anchored Google editor comments when Drive stores custom anchors that editors display as unanchored.
- Treating fine-grained synchronization as a prerequisite for the first stable create/read/import/replace contract.

## Official Platform References

The capability comparison in this roadmap uses current official Google documentation:

- [Google Docs tabs](https://developers.google.com/workspace/docs/api/how-tos/tabs)
- [Google Docs batchUpdate and WriteControl](https://developers.google.com/workspace/docs/api/reference/rest/v1/documents/batchUpdate)
- [Google Sheets batch updates](https://developers.google.com/workspace/sheets/api/guides/batchupdate)
- [Google Sheets values](https://developers.google.com/workspace/sheets/api/guides/values)
- [Google Sheets developer metadata](https://developers.google.com/workspace/sheets/api/guides/metadata)
- [Google Sheets smart chips](https://developers.google.com/workspace/sheets/api/guides/chips)
- [Google Slides API overview](https://developers.google.com/workspace/slides/api/guides/overview)
- [Google Drive uploads and Office conversion](https://developers.google.com/workspace/drive/api/guides/manage-uploads)
- [Google Drive downloads and exports](https://developers.google.com/workspace/drive/api/guides/manage-downloads)
- [Google Drive comments](https://developers.google.com/workspace/drive/api/guides/manage-comments)
- [Google Drive sharing and capabilities](https://developers.google.com/workspace/drive/api/guides/manage-sharing)
- [Google Drive shared drives](https://developers.google.com/workspace/drive/api/guides/manage-shareddrives)
- [Google Drive change tracking](https://developers.google.com/workspace/drive/api/guides/about-changes)
- [Google Drive error handling](https://developers.google.com/workspace/drive/api/guides/handle-errors)
- [Google Workspace authentication overview](https://developers.google.com/workspace/guides/auth-overview)
- [OAuth for installed applications](https://developers.google.com/identity/protocols/oauth2/native-app)
- [OAuth security practices](https://developers.google.com/identity/protocols/oauth2/resources/best-practices)
