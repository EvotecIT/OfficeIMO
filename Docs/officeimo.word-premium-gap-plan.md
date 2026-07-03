# OfficeIMO.Word Premium Gap Plan

This plan turns the remaining non-PDF, non-legacy-DOC Word gaps into small implementation slices that can be handled one by one. It is scoped to `OfficeIMO.Word` and shared OfficeIMO package infrastructure where the capability is reusable across Word, Excel, and PowerPoint.

Out of scope for this plan:

- Word-to-PDF fidelity and PDF proof gates.
- Legacy binary `.doc` read/write support.
- Broad conversion-roadmap work that belongs to HTML, Markdown, RTF, or Reader packages unless it directly proves the Word capability being added here.

## Current Focus

The next market-readiness push should make OfficeIMO.Word credible for four mature Word automation workflows:

1. Digital signature inspection, policy, and eventually signing/validation where it can be done honestly.
2. Review and redline workflows built on comments, revisions, and deterministic review metadata.
3. Document comparison that produces machine-readable findings and optional redline documents.
4. Field evaluation and refresh for common report/document fields, including TOC, REF, PAGEREF-style references, document properties, and simple formula-like fields.

Template/mail-merge polish remains important, but it should follow the review/diff/field foundations unless a user scenario needs it earlier.

## Workstream 0: Shared Planning And Proof Groundwork

Goal: avoid five separate feature implementations with five different report models.

- [x] Define a shared Word proof fixture folder under `OfficeIMO.TestAssets/Documents/Word/PremiumGaps`.
- [x] Add `OfficeIMO.TestAssets/Documents/Word/PremiumGaps/premium-gap-fixtures.xml` to record the source document, feature family, expected behavior, fixture status, evidence, and validation command for each priority gap.
- [x] Add a contract test that keeps the manifest aligned with this plan and prevents future workstreams from drifting into separate proof models.
- [x] Extend `WordFeatureReport` only where it improves preflight for these workstreams; do not turn it into a second document model.
- [x] Keep reusable package-level behavior outside Word when the file-format feature is not Word-specific. Digital-signature package inspection and the bounded package-signing adapter now live in `OfficeIMO.Shared`, with `OfficeIMO.Word` mapping them to Word-facing public models. Cryptographic trust validation still remains future shared-owner or adapter work.
- [x] Keep public examples on normal installed-package entry points: `using OfficeIMO.Word;`, `WordDocument.Load(...)`, `document.InspectFeatures()`, `document.Save(...)`.
- [x] Add one compatibility-matrix row update per completed workstream so users can tell what is authorable, editable, preserved, validated, or intentionally unsupported.

Exit criteria:

- A future agent can pick a workstream and know the source owner, tests, docs, and unsupported boundary before writing code.
- New tests protect current contracts rather than only proving old gaps are gone.

## Workstream 1: Digital Signatures

Goal: make signed-document handling explicit and safe before attempting package signing.

### 1A. Signature Inspection And Mutation Policy

- [x] Inventory current signature surfaces in Open XML packages: `_xmlsignatures`, `DigitalSignatureOriginPart`, package relationships, extended application signature metadata, and macro-signing boundaries.
- [x] Add a public inspection model, `WordSignatureInfo`, with stable fields for presence, part counts, relationship ids, signer metadata when available, digest/signature algorithms when parseable, and unsupported/unknown details.
- [x] Extend `WordFeatureReport` so signed documents show signature metadata presence and validation-unsupported status.
- [x] Add a save-time policy option for signed documents. The default blocks saves that may invalidate signatures; callers can explicitly continue with `WordSaveOptions.SignedDocumentPolicy = WordSignedDocumentSavePolicy.AllowSignatureInvalidation`.
- [x] Add tests proving signature package metadata is preserved by no-op load/save where possible.
- [x] Add tests proving mutating saves surface the policy warning/blocker.

Exit criteria:

- Users can now detect signature package metadata before editing, and save attempts are blocked by default unless the caller explicitly accepts signature invalidation.
- OfficeIMO does not silently imply signature validity after mutation.

### 1B. Signature Validation Feasibility

- [x] Keep first-pass dependency-free validation in `OfficeIMO.Word` as structural signature preflight. Cryptographic validation and package signing remain candidates for a shared package-level owner or optional adapter.
- [x] Validate signed-package structural parsing across `netstandard2.0`, `net8.0`, `net10.0`, and Windows-only `net472` where applicable.
- [x] Add `WordSignatureValidationReport` to separate package structure, XML signature parsing, signed package-part reference coverage, cryptographic validation, certificate-chain trust, revocation, timestamp status, and unsupported validation gaps.
- [x] Add first-pass signed-reference inventory for XML DSig `Reference` entries, including reference digest method/value metadata, package-part target URI resolution, missing digest-value diagnostics, and missing package-part diagnostics.
- [x] Add bounded signed package-part digest verification for simple transform-free package-part references with supported SHA digest methods, while keeping XML DSig transform-aware OPC relationship/content-type reference validation open.
- [x] Add bounded timestamp metadata readback for OPC `SignatureTime` and XAdES `SigningTime` declarations, while keeping timestamp authority/trust validation open.
- [x] Keep chain trust and revocation checks explicitly `NotChecked`; do not make network or machine-certificate behavior implicit.
- [x] Add fixture-backed tests for parseable synthetic signature metadata, missing signed package-part references, malformed signature XML, application-metadata-only signatures, and unsigned documents.
- [x] Add real signed fixture validation from a sanitized OPC package-services signed DOCX created with a temporary OfficeIMO fixture certificate.

Exit criteria:

- Validation claims are now precise for structural parsing: package structure, XML signature reference digest metadata, signer certificate subject readback from package certificate parts, signed package-part target references, timestamp declaration readback, and bounded transform-free package-part digest verification can pass, fail, or report unsupported independently, with proof from a sanitized real signed DOCX fixture and focused synthetic digest/timestamp fixtures. Cryptographic signature validity, transform-aware OPC digest validation, certificate-chain trust, revocation, and timestamp authority validation remain explicit open states until a real validation owner is added.

### 1C. Package Signing

- [x] Start only after inspection and validation boundaries are stable.
- [x] Prefer platform/package APIs if they can sign Open Packaging Convention packages correctly across target frameworks. The first adapter uses `System.IO.Packaging.PackageDigitalSignatureManager` where it is available.
- [x] If signing requires platform-specific APIs or external tools, expose that as an optional adapter rather than forcing the core Word package to own broad certificate-policy behavior. `WordDocument.TrySignPackage(...)` reports unsupported targets, while `SignPackage(...)` fails closed when signing or structural readback cannot be proven.
- [x] Add first-pass signing options for an explicit `X509Certificate2`, selected package-part URIs, package/part relationship inclusion, hash algorithm, optional signature id, and fail-closed behavior when signing cannot be proven.
- [x] Add bounded certificate-store discovery for package signing by thumbprint with explicit store name/location, invalid-certificate inclusion, and private-key requirements.
- [x] Add round-trip tests proving the signed output is structurally valid and validation can read the new signature on the supported Windows .NET Framework adapter.

Exit criteria:

- `OfficeIMO.Word` can sign packages with validator-backed evidence on the supported Windows .NET Framework adapter, using either an explicit certificate or a thumbprint-resolved certificate from an explicit local certificate store, and explicitly reports unsupported targets instead of pretending cross-platform signing exists. Cryptographic signature validation, transform-aware OPC digest validation, certificate-chain trust, revocation, timestamp authority validation/timestamping services, macro-project signing, and cross-platform signing remain open.

## Workstream 2: Review And Redline Workflows

Goal: make OfficeIMO.Word useful for document review automation, not only for toggling track changes.

### 2A. Review Metadata Read Model

- [x] Add a review read model for classic comments, modern/threaded comments, replies, resolved state, authors, initials, dates, ranges, and target text where the document exposes it.
- [x] Add a revision read model for insertions, deletions, moves where parseable, formatting changes where parseable, author, date, affected text, and location.
- [x] Keep raw Open XML escape hatches internal; expose stable OfficeIMO objects for normal use.
- [x] Add feature-report details for comment/revision counts and unsupported modern metadata.
- [x] Add fixture tests for classic comments, reply threads, resolved comments, inserted text, deleted text, and author-filtered revisions.
- [x] Add imported-shape tests proving comment/revision readback inside content controls and text boxes.
- [x] Add imported-shape tests proving move ranges and run/table/cell formatting-change readback.
- [x] Add generated/imported-shape tests proving comment/revision readback and report locations in headers, footers, footnotes, and endnotes.
- [x] Add report proof for newer commentsExtensible metadata as preserved-but-not-parsed unsupported review metadata.
- [x] Add imported-shape proof that comment reply/resolution metadata is matched by comment paragraph id rather than commentsEx storage order.
- [x] Add a sanitized Word-authored review fixture with one real Word comment and real inserted/deleted revision markup.
- [x] Add a second sanitized Word-authored review fixture with real Word move ranges plus run/table/table-row/table-grid/cell formatting revision markup.
- [x] Add a third sanitized Word-authored review fixture with real Word paragraph and section formatting revision markup.
- [x] Add opt-in live Word COM validation that creates a Word-authored body/table comment corpus plus inserted/deleted tracked revisions when `OFFICEIMO_RUN_WORD_COM_VALIDATION=1`.
- [x] Add a sanitized Word COM-authored fixture with body and table comments plus inserted/deleted tracked revisions, so the same real Word-authored shape is covered in normal tests without requiring Word automation at runtime.
- [x] Add a sanitized Word COM-authored related-part revision fixture with header, footer, footnote, and endnote inserted/deleted tracked revisions.
- [x] Add a sanitized Word COM-authored threaded/resolved comment fixture that proves real Word reply/resolution metadata and cross-paragraph comment target readback.
- [x] Add a stable imported related-part comment/revision corpus fixture for header, footer, footnote, and endnote comment target locations, replacing the previous generate-at-test-runtime proof.
- [x] Record the Word COM authoring boundary: Microsoft Word COM refuses to add comments outside the main story, so related-part comment coverage is imported-package interoperability proof rather than a Word-authored related-part comment fixture.

Exit criteria:

- A caller can load a reviewed DOCX and produce a deterministic `WordReviewInfo` summary of comments and tracked changes, including imported-style content-control/text-box review locations, header/footer/note review locations, move ranges, run/table/table-row/table-grid/cell formatting changes, paraId-based comment reply/resolution metadata matching when commentsEx storage order differs, preserved-but-not-parsed commentsExtensible metadata diagnostics, stable imported related-part comment/revision corpus proof for header/footer/footnote/endnote comment targets, first-pass proof against sanitized Word-authored fixtures for comments, inserted/deleted revisions, move ranges, run/table/table-row/table-grid/cell formatting revisions, paragraph formatting revisions, section formatting revisions, a sanitized Word COM-authored body/table comment plus inserted/deleted revision corpus, sanitized Word COM-authored threaded/resolved comments with cross-paragraph targets, and sanitized Word COM-authored header/footer/footnote/endnote inserted/deleted revisions that run without Word automation at test time. Broader real Word-authored imported review corpus coverage remains open, and Word-authored related-part comments are documented as unavailable through Word COM authoring.

### 2B. Review Operations

- [x] Keep existing `AcceptRevisions` and `RejectRevisions` behavior, then expand through the read model rather than separate ad hoc helpers.
- [x] Add scoped operations for accept/reject by author, date range, revision type, paragraph scope, table scope, and explicit revision id when available.
- [x] Add move-range accept/reject proof for imported-style `MoveFrom` and `MoveTo` revisions.
- [x] Add scoped operation coverage for headers, footers, footnotes, and endnotes.
- [x] Add scoped operation coverage for content controls and text boxes.
- [x] Add scoped operation coverage for sanitized Word-authored inserted/deleted revisions and move ranges.
- [x] Add scoped operation coverage for the sanitized Word COM-authored body/table comment and inserted/deleted revision fixture.
- [x] Add scoped operation coverage for the sanitized Word COM-authored related-part revision fixture.
- [x] Add formatting-change accept/reject semantics for property-container revisions where prior properties are stored in the change marker, including run, paragraph, table, table-row, table-cell, and section formatting proof.
- [x] Add broader scoped operation coverage for additional imported Word-authored revision shapes, including sanitized Word-authored paragraph and section formatting revisions.
- [x] Add comment operations for replies, mark resolved/unresolved, delete thread, and delete only one comment.
- [x] Add comment report extraction as part of the structured `WordReviewReport` output, including grouped parent/reply `CommentThreads` plus JSON and Markdown output.
- [x] Preserve unsupported modern comment metadata when operations do not intentionally edit it.
- [x] Add tests proving comment operations update only the intended review items.
- [x] Add tests proving comment operations update the intended thread when imported commentsEx metadata is out of storage order.
- [x] Add tests proving scoped revision operations update only the intended review items.

Exit criteria:

- Comment operations, grouped comment-thread reporting, and first-pass scoped revision operations are precise enough for common thread workflows and common accept/reject automation across body, table, header, footer, footnote, endnote, content-control, and text-box locations, including imported commentsEx ordering that can differ from comment storage order, sanitized Word-authored inserted/deleted/move-range revision markup, sanitized Word-authored paragraph/section formatting revision markup, sanitized Word COM-authored body/table review markup, sanitized Word COM-authored related-part revision markup, and property-formatting revision accept/reject semantics for run, paragraph, table, table-row, table-cell, and section changes. Broader imported review corpus coverage and additional advanced Word-authored revision shapes remain open.

### 2C. Redline Report Output

- [x] Add a structured `WordReviewReport` model that can be serialized to JSON and Markdown.
- [x] Include sections for comments, revisions, unresolved threads, accepted/rejected actions, unsupported review metadata, and document locations.
- [x] Add examples that produce a review report from a document without requiring Word or Office automation.
- [x] Keep visual redline rendering separate from the report model; the report is the contract.
- [x] Add report fixtures for imported-style review metadata inside content controls and text boxes.
- [x] Add report fixtures for imported-style move ranges and run/table/cell formatting changes.
- [x] Add report fixtures for generated/imported-shape review metadata in headers, footers, footnotes, and endnotes.
- [x] Add report fixture proof for imported commentsEx order that differs from comment storage order.
- [x] Add report fixture proof for a sanitized Word-authored comment and inserted/deleted revisions.
- [x] Add report fixture proof for sanitized Word-authored move ranges and run/table/table-row/table-grid/cell formatting revisions.
- [x] Add report fixture proof for sanitized Word-authored paragraph and section formatting revisions.
- [x] Add opt-in live Word COM report proof for Word-authored body/table comments and inserted/deleted revisions when `OFFICEIMO_RUN_WORD_COM_VALIDATION=1`.
- [x] Add normal report fixture proof for the sanitized Word COM-authored body/table comment and inserted/deleted revision corpus.
- [x] Add normal report fixture proof for the sanitized Word COM-authored related-part revision corpus.
- [x] Add broader report fixture proof for sanitized Word COM-authored threaded/resolved comments.
- [x] Add normal report fixture proof for stable imported header/footer/footnote/endnote comment target locations.

Exit criteria:

- A CI job or service can fail, warn, or attach JSON/Markdown report artifacts based on review metadata, including content-control/text-box review locations, header/footer/note review locations, move ranges, run/table/table-row/table-grid/cell formatting changes, commentsEx metadata matched by paragraph id, stable imported header/footer/footnote/endnote comment target locations, sanitized Word-authored fixtures for comments, inserted/deleted revisions, move ranges, run/table/table-row/table-grid/cell formatting revisions, paragraph formatting revisions, section formatting revisions, sanitized Word COM-authored body/table comments plus inserted/deleted revisions, sanitized Word COM-authored threaded/resolved comments, and sanitized Word COM-authored header/footer/footnote/endnote inserted/deleted revisions. Broader real Word-authored imported review corpus coverage remains open; public report examples are covered by the `--word-review-reports` examples workflow.

## Workstream 3: Document Comparison

Goal: turn the existing comparer into a complete document-diff workflow.

### 3A. Structured Diff Depth

- [x] Extend `WordDocumentComparer.CompareStructure(...)` from paragraph/table/image findings into first-pass run-level text and formatting differences.
- [x] Add first-pass field and content-control comparisons for field instruction/result changes and content-control alias/tag/data-binding/text changes.
- [x] Add first-pass bookmark, hyperlink, and list comparison findings with stable feature locations.
- [x] Add first-pass comment and revision comparison findings from structured review metadata.
- [x] Add richer `DetailedLocation` paths for feature findings without breaking existing short `Location` values.
- [x] Add generated-id and volatile-metadata comparison switches for noisy feature metadata.
- [x] Add additional option-aware comparisons for richer feature filters. `IncludedScopes` and `ExcludedScopes` let callers include or remove findings by `WordComparisonScope`; style-id switches are covered for paragraph and run style ids, and bounded effective-formatting comparison now covers document defaults and based-on style chains.
- [x] Include stable paths or locations so findings can be mapped back to document sections, paragraphs, tables, rows, cells, and runs.
- [x] Add first-pass comparison options for ignoring whitespace/case and toggling run formatting, fields, content controls, bookmarks, hyperlinks, lists, comments, revisions, images, and block-order findings.
- [x] Add comparison options for generated ids and volatile metadata.
- [x] Add comparison options for richer feature filters. Scope include/exclude filters, paragraph/run style-id switches, and review metadata subfamily filters now exist.
- [x] Add bounded effective-formatting comparison for document defaults, based-on paragraph/character style chains, paragraph-style run properties, and direct paragraph/run properties through `CompareEffectiveFormatting`.
- [x] Add tests for deterministic ordering and stable findings across repeated runs for paragraph/table/image/run findings.
- [x] Add artifact-backed tests for first-pass field and content-control findings.
- [x] Add artifact-backed tests for first-pass bookmark, hyperlink, list, and comparison-option findings.
- [x] Add artifact-backed tests for first-pass comment/revision findings and option switches.
- [x] Add deterministic-order proof for richer field/content-control edge cases, including repeated comparisons of body, table, header, and content-control-hosted field findings.
- [x] Add deterministic-order proof for imported-style review metadata where comments-part storage order differs from target document order.
- [x] Add artifact-backed comparison tests for generated/imported-shape review metadata in headers, footers, footnotes, and endnotes.
- [x] Add artifact-backed comparison tests for sanitized Word-authored comment, insertion/deletion, move-range, run/table/table-row/table-grid/cell formatting revision metadata, and Word COM-authored related-part inserted/deleted revisions.
- [x] Add static-fixture comparison proof for imported header/footer/footnote/endnote comment targets and inserted revisions from `imported-related-part-comments-revisions.docx`, including JSON/Markdown report coverage.

Exit criteria:

- Two DOCX files can be compared into a stable machine-readable diff without relying on Word automation. First-pass run text/formatting, paragraph/run style-id findings, bounded effective-formatting findings from document defaults and based-on style chains, field/content-control, bookmark, hyperlink, list, comment, revision, header/footer/note review metadata locations, static imported related-part comment/revision fixture proof, sanitized Word-authored review metadata proof including Word COM-authored related-part revision proof, detailed feature locations, generated-id/volatile-metadata filtering, scope include/exclude filters, review metadata subfamily filters, and comparison-option findings now exist; full Word-compatible style cascade parity, toggle-property semantics, theme resolution, and layout-derived effective formatting remain open.

### 3B. Diff Report Formats

- [x] Add JSON and Markdown serializers for comparison results.
- [x] Add a short text summary for CLI/CI wrappers with finding counts by scope and change kind.
- [x] Add examples under `OfficeIMO.Examples/Word` that compare two documents and save the report. The examples runner exposes this through `--word-comparison-reports` and now proves body/header/footer/note paragraph text, run-formatting, field, content-control, and comment metadata findings through the public workflow.
- [x] Keep report generation separate from the comparer so wrappers can choose their own output. `WordComparisonReportWriter` owns JSON, Markdown, and text-summary output, while `WordComparisonResult.ToJson()`, `ToMarkdown()`, and `ToTextSummary()` remain compatibility wrappers.

Exit criteria:

- A human can review the diff report, and automation can consume the same result object. JSON, Markdown, compact text summaries, dedicated report-writer APIs, compatibility wrappers, and public examples for body/header/footer/note text, run-formatting, field, content-control, and comment metadata findings now exist.

### 3C. Redline Document Generation

- [x] Add an optional redline document generator after structured diff is stable.
- [x] Use tracked insertions/deletions for text-bearing structured findings in a generated review artifact.
- [x] Add options for author name, timestamp, summary section, findings table, and whether text findings should become tracked revisions.
- [x] Add Open XML validation coverage for generated redline documents.
- [x] Add first-pass in-place target-document redline generation for body/header/footer/note paragraph text changes, body/header/footer/note run text insertions/deletions, content-control text changes including table-contained run/block controls, native table cell/row SDT variants, nested descendant run controls, and simple text-box-hosted run/block controls, inserted/deleted/changed Drawing and VML image runs, body table-cell text changes/insertions/deletions, body table-row insertions/deletions, body whole-table insertions/deletions, and inserted/deleted nested whole-table redlines in matching parent cells, preserving the target document shell instead of only producing a standalone review artifact.
- [x] Add bounded note-contained table redline coverage for footnote table-cell changes, endnote inserted table rows, and deleted whole footnote tables, reusing the structured table order across related parts.
- [x] Add bounded header/footer table redline coverage for header table-cell changes, footer inserted table rows, and deleted whole footer tables, reusing the structured table order across related parts.
- [x] Add bounded nested table-cell modification redline coverage for an existing nested table inside a matching parent cell.
- [ ] Keep full Word-style in-place redline generation open for advanced nested-table move semantics and broader nested-table layout/modification cases, advanced note shapes beyond plain paragraphs/runs and bounded note-contained table cases, advanced text-box content-control shapes beyond simple run/block SDTs, advanced nested content-control container-wide semantics, full Word-compatible effective formatting semantics beyond the bounded style-chain comparer, advanced header/footer shapes beyond plain paragraphs/runs and bounded table cases, advanced image anchoring/layout semantics, and more complex table move/merge scenarios.
- [x] Add richer policy options for formatting-only findings, comment/revision diffs, and report-only feature findings.
- [x] Add examples under `OfficeIMO.Examples/Word` that compare two documents and save redline outputs. The examples runner exposes this through `--word-comparison-reports` and validates both the standalone review artifact and first-pass in-place target redline while the comparison inputs include body/header/footer/note text, table-cell, table-row, whole-table, run-formatting, field, content-control, and comment metadata differences.

Exit criteria:

- OfficeIMO can now produce a DOCX review artifact from two input documents and a first-pass in-place target-layout redline for supported body/header/footer/note paragraph and run text changes, bounded run-formatting changes as tracked run-property revisions, content-control text changes including table-contained run/block controls, native table cell/row SDT variants, nested descendant run controls, and simple text-box-hosted run/block controls, inserted/deleted/changed Drawing and VML image runs, body table-cell text changes, body table-row insertions/deletions, body whole-table insertions/deletions, bounded note-contained table changes, bounded header/footer table changes, and inserted/deleted nested whole-table plus bounded nested table-cell modification redlines in matching parent cells, with explicit limitations and redline policies for feature, review, and formatting-only findings. The public workflow exercises body/header/footer/note text, table-cell, table-row, whole-table, run-formatting, field, content-control, and comment metadata differences. Full Word-style in-place visual redline generation across advanced nested-table move semantics and broader nested-table layout/modification cases, advanced note shapes beyond plain paragraphs/runs and bounded note-contained table cases, advanced text-box content-control shapes beyond simple run/block SDTs, advanced nested content-control container-wide semantics, full Word-compatible effective formatting semantics beyond the bounded style-chain comparer, advanced header/footer shapes beyond plain paragraphs/runs and bounded table cases, advanced image anchoring/layout semantics, and more complex table move/merge scenarios remains open.

## Workstream 4: Field Evaluation And Refresh

Goal: make common generated-document fields useful without relying entirely on Word opening the document.

### 4A. Field Inventory And Parser

- [x] Inventory simple and complex fields in body, headers, footers, tables, footnotes, endnotes, text boxes, and content controls.
- [x] Add `WordDocument.InspectFields()` with `WordFieldInfo` as the shared readback model for simple and complex field instructions.
- [x] Preserve original instruction text, parsed field type, instructions, switches, format switches, result text, dirty state, locked state, part location, container flags, nesting level, and unsupported parser diagnostics, including recognized-but-unsupported `\*` named format switches and unsupported non-deterministic `\#` numeric picture switches without losing known field identity.
- [x] Add tests for nested fields, fields split across runs, unsupported field instructions, notes, headers/footers, tables, content controls, and text boxes.

Exit criteria:

- Field readback is reliable enough to drive refresh and comparison through `WordDocument.InspectFields()`.

### 4B. Supported Evaluators

- [x] Start with deterministic document metadata fields: `AUTHOR`, `TITLE`, `SUBJECT`, `KEYWORDS`, `COMMENTS`, `CREATEDATE`, `SAVEDATE`, `FILENAME`, and custom document properties.
- [x] Add `WordDocument.UpdateFieldsAndGetReport()` with per-field `Updated`, `Skipped`, `Unsupported`, and `ParseError` diagnostics while keeping `UpdateFields()` as the compatibility wrapper.
- [x] Add bounded `DATE`, `TIME`, and `PRINTDATE` refresh, plus custom `\@` date/time format switches for `DATE`, `TIME`, `CREATEDATE`, `SAVEDATE`, and `PRINTDATE`, using `WordFieldUpdateOptions.CurrentDateTime` when callers need reproducible current date/time fields.
- [x] Add bounded metadata refresh for direct `REVNUM`, `INFO` built-in property fields, and built-in `DOCPROPERTY` aliases such as category, version, revision, and last-printed metadata.
- [x] Add bounded document-statistics refresh for `NUMWORDS` and `NUMCHARS` using OfficeIMO document statistics or package extended properties, including deterministic numeric picture switches.
- [x] Add bounded saved-package `FILESIZE` refresh using the document backing file size, including Word-style bytes, rounded decimal `\k` kilobytes, rounded decimal `\m` megabytes, and deterministic numeric picture switches.
- [x] Add bounded `DOCVARIABLE` refresh from OfficeIMO document variables, including quoted/unquoted field names, case-insensitive lookup, per-field skipped diagnostics for missing variables, and saved-document readback.
- [x] Add bounded `SECTION` and `SECTIONPAGES` refresh for body fields using OfficeIMO section order and explicit page-break counts within the containing body section, including deterministic numeric picture and supported general numeric format switches.
- [x] Add references where the target is already modeled: `REF` bookmark text and `PAGEREF` body bookmark page estimates based on OfficeIMO page-break order.
- [x] Add first-pass generated caption sequence refresh for `SEQ` fields and bookmarked caption `REF` targets, including explicit next (`\n`), reset (`\r`), repeat-current (`\c`), and heading-level reset (`\s`) switches.
- [x] Add bounded cross-reference formatting switches where OfficeIMO can calculate the result deterministically: REF text casing (`\* Upper`, `\* Lower`, `\* FirstCap`, `\* Caps`), PAGEREF numeric formats (`\* Arabic`, `\* Roman`, `\* roman`, `\* Ordinal`, `\* Alphabetical`, `\* ALPHABETICAL`, `\* Hex`, `\* CardText`, `\* OrdText`, `\* DollarText`), and body `PAGEREF` numeric picture switches.
- [x] Add bounded REF list-number cross-reference switches for bookmarks in directly numbered body paragraphs, paragraph styles with `numPr` including simple `basedOn` inheritance, and numbering-level `pStyle` paragraph-style links: `\n` current-level number, `\w` full context, `\r` relative context, and `\t` text suppression, with explicit unsupported diagnostics for bullets and unsupported numbering formats.
- [x] Add page/count fields only where OfficeIMO can calculate them honestly, including bounded `PAGE`/`NUMPAGES` numeric picture switches; otherwise keep `UpdateFieldsOnOpen` and diagnostics.
- [x] Add formula-like fields for bounded arithmetic expressions, percent numeric literals, postfix percent on grouped/function expression results, deterministic numeric functions (`SUM`, `AVERAGE`, `MIN`, `MAX`, `PRODUCT`, `COUNT`, `IF`, `AND`, `OR`, `NOT`, `TRUE`, `FALSE`, `MOD`, `SIGN`, `ABS`, `INT`, `DEFINED`, `ROUND`) with comma or semicolon argument separators, bounded negative-place `ROUND`, short-circuit `IF` branch evaluation, short-circuit `AND`/`OR` logical evaluation, comparison operators (`=`, `<>`, `<`, `<=`, `>`, `>=`), deterministic numeric picture switches for decimal/thousands/percent output, literal text suffixes/prefixes, backslash-escaped literal text, visual color tags, positive/negative/zero sections, explicit numeric condition sections, and layout-neutral Word fill-token normalization, and simple numeric table references over plain numeric and percent-valued cells, including `SUM`/`AVERAGE`/`MIN`/`MAX`/`PRODUCT`/`COUNT` over `ABOVE`, `BELOW`, `LEFT`, and `RIGHT`, explicit `A1`-style cell references, explicit rectangular ranges such as `B1:C2`, explicit `RnCn` cell references such as `R1C1`, explicit `RnCn` ranges such as `R1C1:R2C2`, visual-column resolution across simple horizontally spanned table cells, filtering of vertical-merge continuation cells, row `gridBefore` offsets, and imported-style split complex formula fields for those supported table/percent scenarios.
- [x] Add bounded nested complex-field refresh semantics for result-hosted nested fields: nested deterministic fields inside unsupported containing fields can refresh in place, while nested fields inside a containing field whose whole result was replaced are skipped with a diagnostic so the containing field result is not corrupted.
- [x] Add bounded nested field-instruction evaluation for complex fields: nested deterministic fields inside a containing formula instruction refresh first, and the containing formula consumes the refreshed nested result while readback preserves the effective instruction text.
- [x] Add bounded literal `QUOTE` field refresh for simple and imported-style complex fields, including deterministic text format switches (`\* Upper`, `\* Lower`, `\* FirstCap`, `\* Caps`), numeric literal formats (`\* Arabic`, `\* Roman`, `\* roman`, `\* Ordinal`, `\* Alphabetical`, `\* ALPHABETICAL`, non-negative `\* Hex`, and non-negative integer `\* CardText`, `\* OrdText`, and `\* DollarText`), and numeric picture switches for numeric quoted literals, while keeping argument-free/container-style `QUOTE` fields plus decimal/localized word and currency word formats unsupported.
- [x] Report recognized-but-unsupported parser diagnostics for known field types as `Unsupported` update results instead of losing the field type behind a generic parse error.
- [x] Add unsupported-field diagnostics instead of silent stale results.

Exit criteria:

- `document.UpdateFields()` refreshes metadata, custom property, document-variable, file-name, bounded current `DATE`/`TIME` fields, property-backed `CREATEDATE`/`SAVEDATE`/`PRINTDATE` fields with bounded `\@` date/time format switches, direct `REVNUM`, bounded `INFO` built-in property fields, built-in `DOCPROPERTY` aliases for category/version/revision/last-printed metadata, document-statistics `NUMWORDS` and `NUMCHARS` fields with bounded numeric picture switches, saved-package `FILESIZE` fields with Word-style bytes, rounded decimal `\k` kilobytes, rounded decimal `\m` megabytes, and bounded numeric picture switches, body page/count including bounded `PAGE`/`NUMPAGES` numeric picture switches, body `SECTION`/`SECTIONPAGES` fields based on OfficeIMO section order and explicit page breaks within the containing section, literal `QUOTE` fields with deterministic text format switches, bounded numeric literal formats, and numeric picture switches for numeric quoted literals, `REF` bookmark text with deterministic text casing, directly numbered or style-linked paragraph references, and header/footer/footnote/endnote bookmark text resolution, `PAGEREF` body page estimates with deterministic numeric formats and bounded numeric picture switches, generated-caption `SEQ` including common numbering/reset/ordinal/hex/CardText/OrdText/DollarText format switches, bounded arithmetic/formula-function fields including percent numeric literals, postfix percent on grouped/function expression results, bounded negative-place `ROUND`, deterministic conditional/logical formulas with comma or semicolon argument separators, short-circuit `IF` branch evaluation, short-circuit `AND`/`OR` logical evaluation, and numeric picture switches with literal text, backslash-escaped literal text, visual color tags, positive/negative/zero sections, explicit numeric condition sections, and layout-neutral Word fill-token normalization, and simple numeric table formulas over plain numeric and percent-valued cells such as `SUM(ABOVE)`, `A1 + B1`, `SUM(B1:C2)`, `SUM(A1; B1; R2C1)`, `R1C1 + R2C2`, and `SUM(R1C1:R2C2)`, including imported-style split complex formula fields for the supported table/percent scenarios, visual-column resolution across simple horizontally spanned table cells, filtering of vertical-merge continuation cells, and row `gridBefore` offsets; `document.UpdateFieldsAndGetReport()` also handles bounded result-hosted nested complex fields without corrupting containing field results, refreshes deterministic nested fields inside containing formula instructions before evaluating the containing field, preserves known field identity for recognized-but-unsupported parser diagnostics, and tells callers what was updated, skipped, unsupported, or malformed. Related-part `PAGEREF`, `PAGE`, `SECTION`, and `SECTIONPAGES` fields still require Word layout context and are reported as skipped. Advanced list-number cross-reference behavior beyond direct, style-`numPr`, and numbering-level `pStyle` body paragraphs, broader `INFO` field aliases beyond package-backed built-in properties, broader `DOCVARIABLE` semantics beyond stored string variables, broader `SECTIONPAGES` layout semantics beyond explicit OfficeIMO page-break estimates, broader `QUOTE` container/general behavior, exact Word-compatible statistics recalculation beyond OfficeIMO statistics/extended-property counts, stream-only or unsaved `FILESIZE` evaluation, localized word and currency word formats, broader Word date/time format grammar beyond invariant custom `\@` patterns, broader complex table-layout formula interpretation, broader nested field-instruction evaluation beyond deterministic formula inputs, broader numeric picture switches over non-deterministic or non-numeric source fields, layout-dependent numeric picture fill expansion, and locale-specific numeric picture tokens remain open.

### 4C. TOC, Index, And List Refresh

- [x] Expand TOC refresh beyond field dirtying where headings and outline levels can be calculated from the document model.
- [x] Support configurable heading ranges and hyperlink/bookmark targets where already available.
- [x] Add generated list-of-figures/list-of-tables refresh for caption paragraphs that use OfficeIMO-refreshable `SEQ Figure` and `SEQ Table` fields, with internal bookmarks and explicit page-break/section-start page estimates.
- [x] Add first-pass index refresh for body `XE "Term"`, `XE "Term:Subterm"`, cross-reference `XE "Term" \t "See Other Term"`, and bounded bookmark page-range `XE "Term" \r "BookmarkName"` fields with deterministic sorting, duplicate page merging, explicit page-break estimates, and skipped-entry diagnostics for malformed or unsupported entries.
- [x] Add nested index subentry refresh for generated `XE "Term:Subterm:Detail"` style paths, preserving first-level `Subterm` compatibility while exposing the full `Subterms`/`Path` in `WordIndexEntry`.
- [x] Add bounded typed-index refresh for imported-style `INDEX \f "A"` and matching `XE \f "A"` entry types, preserving the regenerated field filter and excluding other entry types without treating them as malformed.
- [x] Add bounded bookmark-scoped index refresh for imported-style `INDEX \b "BookmarkName"` fields, filtering visible entries to `XE` fields inside the named paragraph-level bookmark scope while preserving the regenerated field filter.
- [x] Add sanitized Word-authored index page-range fixture coverage for raw complex `INDEX` output with `XE \r` bookmark ranges, including Word's body-level bookmark-end shape.
- [x] Add bounded imported-style index separator refresh for `INDEX \e`, `\l`, `\g`, and `\k`, preserving regenerated field switches and matching visible/report page-reference text, including a sanitized Word-authored raw complex-field fixture.
- [x] Add imported index refresh for Word-authored `XE` entries inside body table cells and body block content controls, including a sanitized raw complex `INDEX` fixture.
- [x] Add imported text-box-hosted index refresh proof for Word-authored `XE` entries inside DrawingML text boxes with VML fallback content.
- [x] Add imported nested table text-box index refresh proof for Word-authored `XE` entries inside DrawingML text boxes with VML fallback content anchored in a table cell.
- [x] Add generated related-part index refresh proof for `XE` entries in headers, footers, footnotes, and endnotes.
- [x] Add bounded imported-style index letter-range filtering for `INDEX \p "A-M"` and the documented special-character form such as `INDEX \p "!--B"`, preserving the regenerated switch while keeping locale-specific collation open.
- [x] Add bounded imported-style index heading separators for `INDEX \h`, including custom Latin heading templates such as `INDEX \h "--A--"`, while keeping locale-specific heading/collation behavior open.
- [x] Add bounded imported-style index column-count readback for `INDEX \c "2"`, preserving the regenerated switch and reporting the requested column count while keeping exact Word-compatible multi-column layout open.
- [x] Add bounded Word-style index concordance marking from two-column concordance documents, inserting safe hidden `XE` fields for body, body table-cell, body block content-control, and text-box-hosted paragraph matches with whole-word matching, duplicate prevention, skipped-entry diagnostics, and compatibility with `RefreshIndex()`.
- [x] Add imported related-part index fixture proof for `XE` entries in headers, footers, footnotes, and endnotes loaded from a saved DOCX and refreshed without Word automation.
- [ ] Keep full Word-compatible index generation open for advanced sorting, advanced page-range layout semantics, exact multi-column layout, broader imported index ecosystems, and locale-specific collation.
- [x] Add tests for added headings, nested heading levels, generated bookmark targets, and existing TOC replacement.
- [x] Add tests for imported Word-generated raw complex-field TOC shapes. Section-boundary page-number behavior is covered for next-page, odd-page, and continuous section breaks.
- [x] Add imported table-cell heading TOC refresh for Word-authored raw complex TOC fixtures, normalizing Word table-cell `_Toc...` anchors to validator-clean OfficeIMO anchors.
- [x] Add imported body content-control heading TOC refresh for Word-authored raw complex TOC fixtures, including validator-clean OfficeIMO anchors inside the content control.
- [x] Add generated text-box-hosted TOC heading refresh proof for Heading-style paragraphs inside VML text boxes.
- [x] Add imported text-box-hosted TOC heading refresh proof for Word-authored Heading-style paragraphs inside DrawingML text boxes with VML fallback content.
- [x] Add imported nested table text-box TOC heading refresh proof for Word-authored Heading-style paragraphs inside DrawingML text boxes with VML fallback content anchored in a table cell.
- [x] Add a sanitized Word-generated raw complex-field TOC fixture that uses direct paragraph outline levels rather than Heading styles.
- [x] Add a sanitized Word-generated TC-field TOC fixture with bounded `\f` identifier and `\l` level refresh support.
- [x] Add a sanitized Word-generated raw simple-field TOC fixture that uses `\t` custom style mappings.
- [x] Add bounded bookmark-scoped TOC refresh for imported-style `TOC \b "BookmarkName"` fields, including a sanitized Word-generated raw complex-field fixture.
- [x] Add bounded TOC page-number suppression for imported-style `TOC \n` and `TOC \n "2-3"` fields, including a sanitized Word-generated raw complex-field fixture.
- [x] Add bounded TOC and caption-list page-number separator refresh for imported-style `TOC \p "..."` fields, preserving the regenerated field switch and visible separator text, including a sanitized Word-generated raw complex-field TOC fixture.
- [x] Add a sanitized Word-authored list-of-figures fixture with Word `SEQ Figure` captions and a raw complex `TOC \c "Figure"` list field.
- [x] Add imported table-cell caption-list refresh for Word-authored `SEQ Figure` captions inside body tables, including a sanitized raw complex-field list-of-figures fixture.
- [x] Add imported content-control caption-list refresh for Word-authored `SEQ Figure` captions inside body block content controls, including a sanitized raw complex-field list-of-figures fixture.
- [x] Add generated text-box-hosted caption-list proof for `SEQ Figure` captions inside VML text boxes.
- [x] Add imported text-box-hosted caption-list proof for Word-authored `SEQ Figure` captions inside DrawingML text boxes with VML fallback content.
- [x] Add imported nested table text-box caption-list proof for Word-authored `SEQ Figure` captions inside DrawingML text boxes with VML fallback content anchored in a table cell.
- [x] Add generated header/footer caption-list proof for `SEQ Figure` captions in related parts, including bookmark creation and list hyperlinks.
- [x] Add generated footnote/endnote caption-list proof for `SEQ Figure` captions in note parts, including bookmark creation and list hyperlinks.
- [x] Add imported related-part caption-list fixture proof for header/footer `SEQ Figure` captions loaded from a saved DOCX and refreshed without Word automation.
- [x] Add imported note-part caption-list fixture proof for footnote/endnote `SEQ Figure` captions loaded from a saved DOCX and refreshed without Word automation.
- [x] Add bounded caption-list page-number suppression for imported-style `TOC \c "Figure" \n` list fields, including a sanitized Word-authored raw complex-field fixture.
- [x] Add bounded caption-list page-number separator refresh for imported-style `TOC \c "Figure" \p "..."` list fields, including a sanitized Word-authored raw complex-field fixture.
- [x] Add a sanitized Word-authored list-of-tables fixture with Word `SEQ Table` captions, an excluded `SEQ Figure`, and a raw complex `TOC \c "Table"` list field.
- [x] Add a sanitized Word-authored list-of-equations fixture with Word `SEQ Equation` captions, an excluded `SEQ Figure`, and a raw complex `TOC \c "Equation"` list field that proves the generic `RefreshCaptionList(...)` path.
- [x] Add a sanitized Word-authored index fixture with raw complex `INDEX` output and Word-generated `XE` entries, including a cross-reference entry.
- [x] Add a sanitized Word-authored index page-range fixture with a raw complex `INDEX` field and `XE \r` entry over a multi-page bookmark.
- [x] Add a sanitized Word-authored text-box index fixture with raw complex `INDEX` output and `XE` entries inside DrawingML/VML text-box content.
- [x] Add a sanitized Word-authored nested table text-box index fixture with raw complex `INDEX` output and `XE` entries inside DrawingML/VML text-box content anchored in a table cell.
- [x] Add generated two-column concordance proof that marks body, body table-cell, body block content-control, and text-box-hosted matches as hidden `XE` fields, rejects unsafe index text, and feeds the existing `RefreshIndex()` path.
- [ ] Continue broader imported TOC, caption/list, and index ecosystem fixtures.

Exit criteria:

- `WordTableOfContent.RefreshEntries()` can regenerate a useful TOC for common generated documents and sanitized Word-generated raw complex-field TOC fixtures for Heading-style body paragraphs, Heading-style body table-cell paragraphs, Heading-style body block content-control paragraphs, generated VML text-box-hosted Heading-style paragraphs, Word-authored DrawingML/VML text-box-hosted Heading-style paragraphs, nested table-cell DrawingML/VML text-box-hosted Heading-style paragraphs, direct-outline-level, bookmark-scoped, page-number-suppressed, page-number-separator, and bounded TC-field source paragraphs, plus a raw simple-field TOC fixture using `\t` custom style mappings, including explicit page-break estimates, section-boundary page estimates, direct paragraph outline levels, validator-clean imported table-cell and body content-control TOC bookmark normalization, generated text-box-hosted TOC bookmark creation, alternate-content text-box heading deduplication, nested table-cell text-box heading deduplication, `TOC \b` bookmark scopes, `TOC \n` page-number suppression, `TOC \p` entry/page separators, `TOC \f` / `TC \f` type filters with `TC \l` levels, and `TOC \t` style-to-level mappings. `WordTableOfContent.RefreshListOfFigures()`, `RefreshListOfTables()`, and `RefreshCaptionList(...)` can regenerate useful generated-caption lists after field refresh and sanitized Word-authored list-of-figures/list-of-tables fixtures plus a Word-authored generic equation-list fixture, including table-cell, body content-control, generated text-box-hosted, Word-authored DrawingML/VML text-box-hosted, nested table-cell DrawingML/VML text-box-hosted, generated header/footer related-part caption paragraphs, generated footnote/endnote note-part caption paragraphs, imported related-part header/footer and note-part list-of-figures fixtures, imported `TOC \n` page-number suppression, and imported `TOC \p` entry/page separators for caption-list entries while keeping estimated page numbers in the report. `WordTableOfContent.RefreshIndex()` can generate a first-pass visible index from body, body table-cell, body block content-control, Word-authored DrawingML/VML text-box `XE` terms, and nested table-cell DrawingML/VML text-box `XE` terms, nested subentry paths, cross-reference `\t` switches, bounded `XE \r` bookmark page ranges, typed-index `INDEX \f` / `XE \f` filters, paragraph-level `INDEX \b` bookmark scopes, bounded imported-style Latin letter-range `INDEX \p` filters, bounded imported-style Latin `INDEX \h` heading separators, bounded imported-style `INDEX \c` column-count readback and switch preservation, and imported-style `INDEX \e` / `\l` / `\g` / `\k` separators, including sanitized Word-authored raw complex `INDEX` fixtures for main/sub/cross-reference entries, bookmark page ranges, custom separators, container-hosted, text-box-hosted, and table text-box-hosted `XE` entries, plus generated header/footer/footnote/endnote related-part `XE` entries. `WordDocument.MarkIndexEntriesFromConcordance(...)` can mark body, body table-cell, body block content-control, and text-box-hosted paragraph matches from a two-column concordance document as hidden `XE` fields before `RefreshIndex()` regenerates the visible index. Full Word-compatible index generation beyond body/table-cell/body-content-control/text-box/table-text-box/header-footer/footnote-endnote `XE` containers and bounded concordance marking, exact multi-column layout, broader imported caption/list ecosystems beyond body/table-cell/body-content-control/text-box/table-text-box and the first header/footer and note-part fixtures, locale-specific index collation/yomi/heading behavior, and deeper imported TOC shapes beyond body, body table-cell, body content-control, text-box-hosted, and table text-box-hosted heading paragraphs remain open.

## Workstream 5: Template And Mail-Merge Polish

Goal: make the existing merge engine easier to trust and easier to sell as a serious open-source document automation feature.

- [x] Publish a scenario matrix for merge fields, conditional blocks, repeated table rows, grouped table rows, repeated body blocks, nested regions, section regions, headers/footers, table cells, and content controls in [`Docs/officeimo.word-template-mail-merge-scenarios.md`](officeimo.word-template-mail-merge-scenarios.md).
- [x] Add section-region fixture proof for conditional and repeating regions that preserve section break and page setup properties.
- [x] Add deeper formatting-preservation tests for split runs, complex fields, nested regions, table-cell content, and content-control replacements.
- [x] Add a `WordTemplatePreflightReport` on top of `InspectTemplate()` for CI/service capability checks, diagnostics, JSON output, and Markdown output.
- [x] Add saved-DOCX artifact proof that `WordTemplatePreflightReport` detects merge fields, conditional markers, and repeated-block markers inside table cells after save/load.
- [x] Add saved-DOCX artifact proof that `WordTemplatePreflightReport` detects merge fields, conditional markers, and repeated-block markers inside headers and footers after save/load.
- [x] Add public examples for invoice, grouped table report, proposal, review letter, header/footer approval package, and form-fill workflows. `OfficeIMO.Examples --word-mail-merge-workflows` now generates all six workflow documents plus preflight/form diagnostics, including grouped table-row preflight proof, header/footer template preflight proof, and JSON/Markdown content-control validation proof artifacts; the market-readiness gallery also includes clean and blocked template preflight proof artifacts.
- [x] Add a sanitized Word-authored multi-section template fixture with Word-created merge fields and a conditional landscape section that can be included or removed.
- [x] Add a sanitized Word-authored content-control form fixture with text, rich text, checkbox, date, dropdown, combo box, picture, and table-cell block content controls that can be validated, filled, saved, and extracted through the public form-map APIs.
- [x] Add preflight diagnostics for unsupported Word-native mail-merge record-control fields (`NEXT`, `NEXTIF`, `SKIPIF`, `MERGEREC`, and `MERGESEQ`) so imported templates do not look safely bindable when they require Word's multi-record merge engine.
- [x] Keep template APIs in `OfficeIMO.Word`; keep PowerShell-friendly wrappers in PSWriteOffice later as a thin surface. The workflow examples call the public OfficeIMO.Word APIs directly and do not add wrapper-side behavior.

Exit criteria:

- Users can validate a template, bind data, generate one or many documents, and receive actionable diagnostics for missing or unsupported template features. The reusable preflight report now covers merge fields, conditional blocks, repeated blocks, saved table-cell-hosted merge/condition/repeating markers, saved header/footer-hosted merge/condition/repeating markers, and unsupported Word-native mail-merge record-control fields such as `NEXTIF` and `SKIPIF`; content-control form-map validation now emits reusable deterministic JSON and Markdown diagnostics; the scenario matrix maps supported template shapes to proof; public workflow examples cover invoice, grouped table report, proposal, review letter, header/footer approval package, and form-fill outputs with generated preflight and JSON/Markdown form-validation artifacts; and artifact tests now cover split-run complex fields, nested regions inside table cells, content-control replacement formatting, section-shaped conditional/repeating regions, a first Word-authored multi-section conditional template, imported-style record-control field diagnostics, and a Word-authored content-control form with rich-text, combo-box, picture, and table-cell block SDTs. Broader imported-template corpus coverage and execution of Word's native multi-record merge-control semantics remain open.

## Workstream 6: Real-World Corpus And Documentation

Goal: turn capability into proof.

- [x] Add a small real-world-style generated corpus for signed documents, reviewed documents, comparison pairs, field-heavy reports, merge templates, and unknown-document feature preflight through `OfficeIMO.Examples --word-market-readiness`.
- [x] Add Open XML validation for generated and mutated DOCX outputs in the market-readiness gallery.
- [x] Add feature-report snapshots for unknown-document preflight in the market-readiness gallery.
- [x] Add dedicated examples folders per major workflow: signatures, review reports, comparison reports, field refresh, and templates. `--word-signature-preflight`, `--word-review-reports`, `--word-comparison-reports`, `--word-mail-merge-workflows`, and the update-field examples now give readers workflow-sized entry points; the market-readiness gallery still generates the combined proof corpus.
- [x] Update `OfficeIMO.Word/COMPATIBILITY.md` after each completed workstream slice so the public matrix names current proof and remaining limits.
- [x] Keep dated investigation notes out of the main docs once the compatibility matrix and examples become the source of truth. The old Word-specific dated review note is no longer referenced from the plan or Word/Excel capability assessment; those docs now point to this plan and the market-readiness docs as the current source of truth.

Exit criteria:

- The public story is not "we have classes"; it is "here is the workflow, here is the proof, and here are the limits." The market-readiness gallery now generates a small proof corpus for premium Word workflows and unknown-document preflight; broader imported real-world fixtures remain open.

## Remaining Gap Size

This branch closes many of the comparison points from the 2026 library review, but the remaining items are not all the same size. Treat them as three groups when deciding what belongs in this PR and what should become follow-up PRs.

### Finish Before This PR

These are release-readiness tasks for the current branch, not new feature work:

- [ ] Rebase or merge current `origin/master`, then resolve any drift in touched Word, shared-signature, fixture, example, and docs files.
- [ ] Review public API names and docs for consistency across signatures, review reports, comparison reports, field reports, template preflight, and TOC/index reports.
- [ ] Run the final focused validation matrix from this document across supported target frameworks, including the Windows-only `net472` signing lane where available.
- [ ] Run at least one full `net8.0` test pass or a documented narrower substitute if the full suite is too expensive for the PR turn.
- [ ] Regenerate or remove any temporary proof artifacts that should not be committed; keep only fixtures, examples, and docs that are part of the proof story.
- [ ] Normalize line endings/formatting noise in touched docs such as `OfficeIMO.Word/COMPATIBILITY.md`.
- [ ] Confirm the PR adds no new runtime NuGet dependency. Test-only dependencies are acceptable, but any runtime dependency must be called out explicitly.
- [ ] Prepare a release-ready PR description that separates completed behavior, intentional limits, validation, and follow-up work.

### Medium Follow-Up Work

These are meaningful but bounded follow-ups. Each should fit in one focused PR if the scope stays narrow and fixture-driven:

- [ ] Add more imported review/redline corpus coverage for real Word-authored documents beyond the sanitized fixtures already present.
- [ ] Expand in-place redline generation for one additional concrete shape at a time, such as a specific nested-table modification pattern, a specific note-contained structure, or a specific text-box content-control shape.
- [ ] Expand comparison effective-formatting coverage for one style semantics slice at a time, such as toggle-property handling or a bounded theme-color resolution path.
- [ ] Add additional imported TOC/caption/index fixtures for real-world documents, especially documents that mix supported containers in one file.
- [ ] Add broader imported-template fixtures for section-heavy templates and content-control-heavy templates that use Office-native authoring patterns.
- [ ] Add more field-refresh diagnostics for known unsupported field switches so callers receive precise skipped/unsupported reasons instead of generic fallback.

### Large Or Separate Owner Work

These are big enough that they should not block this PR unless the product decision is to chase full Word parity before release:

- [ ] Real cryptographic signature validation, including XML DSig validation, transform-aware OPC digest validation, certificate-chain trust, revocation, and timestamp-authority validation. This likely belongs in a shared package-signature owner or optional adapter, not only in `OfficeIMO.Word`.
- [ ] Cross-platform package signing. The current Windows/.NET Framework adapter is bounded; portable signing needs a real OPC-signing implementation or a vetted external adapter.
- [ ] Macro-project signing. This is a distinct VBA/project-signature domain and should stay separate from document package signatures.
- [ ] Full Word-style in-place redline generation for advanced nested-table moves/layout changes, advanced note shapes, advanced text-box SDTs, complex content-control semantics, advanced image anchoring/layout, and table move/merge semantics. This is effectively a Word-layout/revision rendering engine, not a small extension to the current structured diff.
- [ ] Full Word-compatible style cascade and layout-derived effective formatting, including theme resolution, toggle properties, and layout-dependent formatting. This needs a dedicated style/layout compatibility track.
- [ ] Layout-dependent related-part `PAGE`, `PAGEREF`, `SECTION`, and `SECTIONPAGES` evaluation. Without Word or a layout engine, these should remain explicit skipped diagnostics.
- [ ] Full Word formula and locale behavior for fields, including locale-specific date, number, word, and currency formats plus complex table-layout interpretation.
- [ ] Full Word-compatible index generation, including advanced sorting, page-range layout semantics, exact multi-column layout, yomi/locale collation, and locale-specific heading behavior.
- [ ] Execution of Word's native multi-record mail-merge control semantics such as `NEXT`, `NEXTIF`, `SKIPIF`, `MERGEREC`, and `MERGESEQ`. Current work should continue to diagnose or explicitly reject these unless OfficeIMO intentionally implements Word's record engine.

Practical read: the current branch is close to a strong premium-gap PR if final validation and cleanup hold. The remaining hard items are mostly full-fidelity Word compatibility, cryptographic trust, layout, locale, or native Word-engine semantics. Those are large roadmap tracks, not obvious missing small patches.

## Suggested Execution Order

1. Workstream 0: shared proof and report shape.
2. Workstream 4A: field inventory/parser, because comparison and refresh both need stable field readback.
3. Workstream 2A: review metadata read model.
4. Workstream 3A: deeper structured diff using the field and review models.
5. Workstream 1A: signature inspection and mutation policy.
6. Workstream 4B/4C: supported field evaluators and TOC refresh.
7. Workstream 2B/2C: review operations and reports.
8. Workstream 3B/3C: diff reports and redline documents.
9. Workstream 1B/1C: signature validation and signing feasibility.
10. Workstream 5/6: template polish, examples, corpus, and compatibility docs.

This order gives early reusable models to later slices, avoids premature signing promises, and lets agents ship visible value before the hard cryptographic and redline-generation work.

## Suggested Validation Slices

Use focused tests while implementing, then broaden before release:

```powershell
dotnet build OfficeIMO.Word\OfficeIMO.Word.csproj -f net8.0
dotnet test OfficeIMO.Word.Tests\OfficeIMO.Word.Tests.csproj -f net8.0 --filter "FullyQualifiedName~Word.Signature|FullyQualifiedName~Word.Review|FullyQualifiedName~Word.Compare|FullyQualifiedName~Word.UpdateFields|FullyQualifiedName~Word.MailMerge"
dotnet test OfficeIMO.Word.Tests\OfficeIMO.Word.Tests.csproj -f net10.0 --filter "FullyQualifiedName~Word.Signature|FullyQualifiedName~Word.Review|FullyQualifiedName~Word.Compare|FullyQualifiedName~Word.UpdateFields|FullyQualifiedName~Word.MailMerge"
```

On Windows release-prep lanes, include `net472` once the touched code is compatible:

```powershell
dotnet test OfficeIMO.Word.Tests\OfficeIMO.Word.Tests.csproj -f net472 --filter "FullyQualifiedName~Word.Signature|FullyQualifiedName~Word.Review|FullyQualifiedName~Word.Compare|FullyQualifiedName~Word.UpdateFields|FullyQualifiedName~Word.MailMerge"
```

For each workstream, add one artifact-level test that opens the generated DOCX with Open XML SDK validation or a package-level assertion. Feature-level tests are useful, but the user contract is the saved document.
