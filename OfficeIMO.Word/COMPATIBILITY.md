# OfficeIMO.Word compatibility

This matrix describes the current Word contract. “Partial” means useful behavior exists with an explicit preservation, rendering, signing, or interoperability boundary. Open Word work is tracked in the repository [roadmap](../Docs/ROADMAP.md).

## Formats

| Format | Current contract | Boundary |
| --- | --- | --- |
| DOCX/DOCM/DOTX/DOTM | Native create, load, inspect, edit, preserve, and save through `WordDocument` | Unknown or preservation-sensitive parts are reported before edit-heavy workflows |
| DOC/DOT | First-party binary import, supported native writer subset, guarded DOC/DOCX conversion, feature reporting, and source-preserving fallbacks | Unsupported binary features remain diagnosed, preserved, visually represented, or blocked according to compatibility policy |

The detailed legacy contract is documented in [DOC/DOCX compatibility](../Docs/officeimo.word.legacy-doc-compatibility.md) and [Word/Excel interoperability](../Docs/officeimo.word-excel-interoperability.md).

## Document capabilities

| Area | Status | Current contract and boundary |
| --- | --- | --- |
| Create/load/save | Supported | File, byte, and caller-owned stream workflows share the normal lifecycle vocabulary and deterministic save behavior |
| Paragraphs, runs, styles, lists, tables, sections | Broad | Code-first authoring and common imported-document editing are supported; Word-exact layout and every producer-specific style cascade are not claimed |
| Headers, footers, notes, fields, variables, bibliography | Broad | Common authoring, readback, editing, and preservation paths are available with structured diagnostics for unsupported shapes |
| Images, drawings, charts, shapes, and SmartArt | Partial | Common images, charts, shapes, and shared Drawing export are supported; advanced anchoring, imported chart mutation, SmartArt editing, grouping, and Word-exact layout remain bounded |
| Templates and mail merge | Broad | Merge fields, conditionals, repeated rows/blocks/sections, content controls, Custom XML binding, validation, batch output, and formatting-preserving replacements are covered by the [scenario matrix](../Docs/officeimo.word-template-mail-merge-scenarios.md) |
| Content controls and forms | Broad for common SDTs | Text, rich text, checkbox, date, dropdown, combo, picture, repeating-section, tag/alias lookup, fill, extraction, and preflight are supported; advanced mapping/binding shapes remain bounded |
| Comments, revisions, and review reports | Broad | Classic comments/replies/resolution, imported review metadata, scoped accept/reject, visible markup, structured JSON/Markdown reports, and supported redline artifacts are available. Sanitized Word-authored and imported-shape proof covers bounded body/nested-table, note, header, footer, text-box, and content-control locations; complete Word-compatible review semantics are not claimed |
| Document comparison | Broad structured comparison | Deterministic paragraph/run/table/image/field/content-control/bookmark/link/list/comment/revision findings and supported in-place or generated redline output are available. Every report carries stable limitation codes for theme/conditional-style/numbering effective formatting and revision-metadata-only move semantics when those shapes are present |
| Fields, TOC, indexes, and lists | Supported deterministic profiles | Parsing, inventory, selected evaluation, TOC/index/caption-list refresh, diagnostics, and switch preservation are available. Each update reports whether it used the invariant document model, an explicit-break page estimate, or requires an external layout engine |
| Macros and embedded payloads | Inspect/preserve/manage | VBA, package, OLE, and ActiveX payloads can be inventoried, hashed, extracted, attached/replaced/removed without execution; source-level VBA and full OLE/ActiveX editing are outside the contract |
| Digital signatures | OPC package signing and validation | Cross-platform OPC XML-signature creation, relationship-transform and canonicalization-aware digests, signature math, signer-chain trust/revocation policy, RFC 3161 timestamp-authority validation, resource budgets, and save invalidation policy are supported. VBA macro-project signing is a separate explicit unsupported capability; existing projects and signatures can be inspected or preserved according to save policy |
| Protection and encryption | Broad | Document protection, encrypted OOXML workflows, package security, and active/external-content policy are available with typed findings |
| Feature inspection and preflight | Supported | `InspectFeatures()`, capability preflight, `Can`, `EnsureCan`, diagnostics, repair hints, and preservation reports route reads, edits, templates, rendering, and save workflows explicitly |

## Field evaluator

`WordDocument.UpdateFieldsAndGetReport()` refreshes the deterministic subset below and reports each field as updated, skipped, unsupported, or malformed. Unsupported fields and switches are preserved rather than silently rewritten.

| Family | Evaluated subset and boundary |
| --- | --- |
| Metadata and properties | File name; custom and built-in document properties; document variables; `REVNUM`; bounded `INFO` aliases; `NUMWORDS`, `NUMCHARS`, and saved-package `FILESIZE` |
| Dates and time | Current `DATE` / `TIME` plus property-backed `CREATEDATE`, `SAVEDATE`, `PRINTDATE`; invariant custom `\@` patterns are supported, while broader locale-dependent Word date grammar remains outside the deterministic subset |
| Page and section values | Body `PAGE` / `NUMPAGES` and `SECTION` / `SECTIONPAGES` using OfficeIMO section order and explicit page-break estimates; related-part fields that need Word layout are skipped with diagnostics |
| Literal and reference fields | Literal `QUOTE`; `REF` bookmark text and supported numbered-paragraph references; body `PAGEREF` estimates; generated-caption `SEQ` with next, reset, repeat-current, heading reset, and supported number formats |
| General and numeric switches | Text casing `\* Upper`, `\* Lower`, `\* FirstCap`, and `\* Caps`; Arabic, Roman/roman, Ordinal, Alphabetical/ALPHABETICAL, non-negative Hex, CardText, OrdText, and DollarText; bounded numeric picture switches and literal text sections |
| Formula functions | `SUM`, `AVERAGE`, `MIN`, `MAX`, `PRODUCT`, `COUNT`, `IF`, `AND`, `OR`, `NOT`, `TRUE`, `FALSE`, `MOD`, `SIGN`, `ABS`, `INT`, `DEFINED`, and `ROUND`; comma and semicolon separators, percent literals/results, comparisons, and short-circuit logical branches are supported |
| Table formulas | Plain numeric and percent-valued cells with `ABOVE`, `BELOW`, `LEFT`, and `RIGHT`, explicit A1 and `RnCn` cells, and rectangular ranges in regular tables; simple horizontal spans, vertical-merge continuations, and `gridBefore` offsets are normalized |
| Nested complex fields | Deterministic nested result fields and deterministic nested formula inputs refresh within bounded containing-field shapes; fields whose containing result was replaced are skipped to avoid corrupting that result |

TOC, caption-list, and index refresh have their own explicit methods and diagnostics. Their reports expose `WordPageNumberBasis.ExplicitBreakEstimate`; field updates expose `DiagnosticCode` and `EvaluationBasis`. Locale-dependent layout, broader complex table geometry, unsaved `FILESIZE`, broader `QUOTE` container behavior, native `LISTNUM`, and unsupported nested instruction shapes remain preserved and diagnosed rather than approximated silently.

## Package signatures

Use `WordDocument.SignPackage(...)` or its non-throwing `TrySignPackage(...)` form with a certificate that has an accessible private key. `WordPackageSigningOptions` selects parts, relationship selectors, digest algorithm, embedded chain certificates, claimed signing time, and input budgets. Load the package and call `ValidateSignatures(WordSignatureValidationOptions)` to apply caller-selected certificate trust, revocation, timestamp-authority, and resource policy.

`WordSigningCapabilities.Package` and `WordSigningCapabilities.MacroProject` deliberately report two different contracts. A valid OPC package signature does not sign VBA source, and preserving a signed VBA project does not create or renew its signature.

## Conversion and rendering

- [Word/HTML support](../Docs/officeimo.word-html-support-matrix.md) states the bidirectional HTML contract.
- Markdown conversion uses the owning Markdown models and reports destination-specific loss.
- PDF conversion uses the Word adapter over the shared PDF/Drawing owners and returns structured fidelity evidence.
- Managed image export produces PNG, JPEG, TIFF, WebP, and SVG through shared Drawing primitives with estimated pagination and explicit diagnostics.

OfficeIMO does not use Microsoft Word as a runtime renderer. Layout-sensitive claims require fixture and visual evidence, and estimated pagination is not presented as Word-exact pagination.

## Security and preservation

`WordLoadOptions.PackageSecurity` applies shared package, part, aggregate-size, compression-ratio, unsafe-name, relationship, macro, embedded-payload, ActiveX, and external-relationship policy before parsing. Secure defaults retain compatible active content within structural limits; untrusted defaults reject active and external content.

Signed inputs are blocked from ordinary save by default. A caller must explicitly accept signature invalidation or use a proven append-only/signature-aware workflow.

## Validation

- [Word tests](../OfficeIMO.Word.Tests)
- [Office interoperability gate](../Build/Test-OfficeInteroperabilityGate.ps1)
- [Word/Excel interoperability guide](../Docs/officeimo.word-excel-interoperability.md)
- [Template and mail-merge scenario matrix](../Docs/officeimo.word-template-mail-merge-scenarios.md)
- [Word/HTML support matrix](../Docs/officeimo.word-html-support-matrix.md)
- [Image export capability matrix](../Docs/officeimo.image-export-capability-matrix.md)
