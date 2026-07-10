# OfficeIMO RTF support matrix

Last reviewed: 2026-07-10

This is the living product contract for RTF read, write, edit, conversion, and ingestion. The detailed evidence and market comparison are in [the 2026-07-10 end-to-end audit](reviews/officeimo.rtf-end-to-end-market-gap-2026-07-10.md).

Status legend:

- `Full`: public API, focused tests, and appropriate fixture evidence exist.
- `Broad`: substantial support exists, but important fidelity or workflow gaps remain.
- `Partial`: usable subset with known degradation.
- `Extractive`: recovers useful content, not round-trip or visual fidelity.
- `Preserved`: source syntax survives the lossless path but is not semantically editable.
- `Gap`: no adequate current implementation or proof.

Conversion class legend:

- `Lossless`: preserves original RTF syntax for untouched content.
- `Semantic`: preserves editable document meaning.
- `Visual`: targets rendered appearance and may lose source structure.
- `Extractive`: recovers useful text, metadata, or provenance.
- `Diagnostic`: reports preserved, flattened, omitted, or blocked content.

## Package ownership and publication

| Area | Owning package | Current package | Contract |
| --- | --- | --- | --- |
| Parser, syntax, semantic model, writer, diagnostics, lossless editor | `OfficeIMO.Rtf` | 0.1.10 | Reusable core; no adapter-specific mapping belongs here. |
| RTF <=> Word | `OfficeIMO.Word.Rtf` | 0.1.10 | Thin mapping to `OfficeIMO.Word`; higher-level workflows stay in Word. |
| RTF <=> HTML | `OfficeIMO.Html` | Shared HTML package | URL/resource policy and semantic/round-trip HTML mapping. |
| RTF <=> Markdown | `OfficeIMO.Rtf.Markdown` | 0.1.8 | Direct `RtfDocument`/`MarkdownDoc` bridge. |
| RTF <=> PDF | `OfficeIMO.Rtf.Pdf` | 0.1.10 | Semantic visual export and extractive import through `OfficeIMO.Pdf`. |
| RTF => Reader chunks | `OfficeIMO.Reader.Rtf` | 0.0.11 | Extraction/chunking only. |

All listed RTF packages target `netstandard2.0`, `net8.0`, and `net10.0`; Windows builds also target `net472`.

## Core engine

| Capability | Syntax/lossless | Semantic model | Writer | Diagnostics | Status and boundary |
| --- | --- | --- | --- | --- | --- |
| Groups, control words/symbols, text, binary payloads | Full | Full for generic structure | Full | Full for syntax errors | Core parsing foundation is strong. |
| Unknown destinations | Full | Preserved only | Full on lossless path | Partial | Unknown ignorable destinations can be skipped semantically without a warning. |
| Fonts and colors | Full | Broad | Broad | Partial | Core tables are rich; theme/table-style families remain gaps. |
| ANSI, PC, PCA, Mac Roman, Unicode escapes | Full | Broad | Broad | Broad | Supports IBM 437/850, Mac Roman, Windows 874 and 1250-1258. |
| East Asian/DBCS text | Full source preservation | Gap | Partial | Partial | Code pages 932/936/949/950 and composite-font semantics are missing. |
| Paragraphs and runs | Full | Full | Full | Broad | Rich character and paragraph formatting is heavily tested. |
| Character effects and borders | Full | Broad | Broad | Broad | Includes underline variants, caps, hidden, offsets, spacing, shading, and borders. |
| Paragraph layout and frames | Full | Broad | Broad | Broad | Alignment, indents, spacing, pagination, direction, tab stops, and frame metadata. |
| Styles | Full | Partial | Partial | Partial | Basic stylesheet entries exist; quick/latent/table styles and theme mapping do not. |
| Lists and numbering | Full | Broad | Broad | Partial | Modern and legacy lists exist; cross-adapter numbering structure is incomplete. |
| Tables | Full | Broad | Broad | Broad | Rows, cells, merges, widths, padding, shading, borders, and flow are modeled. |
| Nested tables | Full source preservation | Gap | Gap | Gap | `nesttableprops`, `nestrow`, `nestcell`, and `itap` are not modeled. |
| Images | Full | Broad | Broad | Broad | PNG/JPEG/DIB/WMF/EMF are modeled; adapter rendering varies. |
| Fields, hyperlinks, bookmarks, form-field data | Full | Broad | Broad | Partial | Visible results and many field forms work; field evaluation belongs in Word. |
| Notes and annotations | Full | Broad | Broad | Broad | Footnotes, endnotes, annotations, and generated references exist. |
| Headers and footers | Full | Broad | Broad | Partial | Core model is present; target adapters simplify some variants. |
| Sections and page setup | Full | Broad | Broad | Broad | Includes columns, numbering, borders, line numbering, and note placement. |
| Revisions and comments | Full | Partial | Partial | Partial | Insert/delete revisions and annotations work; move revisions are missing. |
| Shapes and objects | Full | Partial | Partial | Partial | Basic semantic objects/shapes exist; drawing fidelity and adapter reporting are incomplete. |
| Metadata, document variables, user properties | Full | Broad | Broad | Broad | Rich document metadata is supported. |
| Custom XML, smart tags, data store | Full source preservation | Gap | Lossless only | Gap | No semantic family contract yet. |
| Outlook HTML-encapsulated RTF | Full source preservation | Gap | Gap | Gap | `fromhtml`, `htmltag`, `htmlrtf`, and `mhtmltag` are not modeled. |
| Index and TOC entry destinations | Full source preservation | Gap | Gap | Gap | No semantic `xe`/`tc` contract. |

## Editing

| Capability | Status | Notes |
| --- | --- | --- |
| Byte/source-preserving round trip | Full | `RtfReadResult.ToRtfLossless` and lossless save APIs preserve untouched syntax. |
| Visible text replacement | Broad | Skips structural destinations and escapes inserted text. Does not provide rich cross-run replacement. |
| Metadata and document settings | Broad | Info, variables, user properties, fonts, colors, styles, page setup, borders, numbering, revisions, files, and XML namespace tables. |
| Append plain paragraph | Full | Available through `RtfLosslessEditor.AppendParagraph`. |
| Insert/remove/move arbitrary blocks | Gap | Needed for structural lossless editing. |
| Replace images and edit table/header/footer content losslessly | Gap | Planned differentiator. |
| Semantic document clone/merge | Gap | Native model operation still missing. |
| Bookmark/range editing | Gap | High-value workflow for document automation. |

## Adapter matrix

| Feature | Word | HTML | Markdown | Reader | PDF export | PDF import |
| --- | --- | --- | --- | --- | --- | --- |
| Plain paragraphs/runs | Full | Full | Full | Full | Full | Extractive |
| Bold/italic/strike/underline | Broad | Broad | Broad | Extractive | Broad | Extractive |
| Font family/size/color/highlight | Broad | Broad | Partial | Extractive | Partial | Extractive |
| Hyperlinks | Broad | Broad but output policy gap | Broad | Full | Broad | Extractive |
| Bookmarks | Broad | Broad | Partial | Extractive | Broad | Gap |
| Headings/styles | Partial | Broad | Broad | Full | Partial | Extractive |
| Ordered/unordered lists | Partial | Broad | Broad | Full | Broad | Extractive |
| Nested lists and restarts | Partial | Broad | Broad | Partial | Partial | Extractive |
| Tables | Broad | Broad | Broad | Full | Broad | Gap |
| Nested tables | Gap | Gap | Gap | Gap | Gap | Gap |
| Merged cells | Broad | Broad | Partial | Partial | Broad | Gap |
| Images | Broad | Broad with profile/resource gap | Partial | Extractive | PNG/JPEG | Gap |
| Footnotes/endnotes | Broad | Broad | RTF->MD gap | Full | Appended semantic notes | Gap |
| Comments/annotations | Broad | Broad | Diagnostic omission | Full | Appended semantic notes | Gap |
| Headers/footers | Broad | Broad | Diagnostic omission | Full | Partial | Gap |
| Sections/page setup | Broad | Broad round-trip metadata | Diagnostic | Extractive | Broad | First-page size only |
| Revisions/tracked changes | Broad subset | Broad subset | Diagnostic | Extractive | Flattened | Gap |
| Shapes/text boxes/objects | Gap or silent omission | Round-trip metadata/text | Diagnostic | Extractive | Plain-text fallback | Gap |
| Unknown destinations | Lossless only; no result report | Lossless source only | Diagnostic for modeled omissions | Core warnings only | Core warnings only | Not applicable |

## Safety and operational contract

| Requirement | Status | Current evidence or gap |
| --- | --- | --- |
| Maximum group depth | Full | `RtfReadOptions.MaxDepth`, default 512. |
| Maximum input size | Adapter-only | Reader enforces `MaxInputBytes`; the core and other adapters do not. |
| Token, text, group, block limits | Gap | No core limits. |
| Binary/image/object limits | Gap | No per-item or total core limits. |
| Cancellation during parse/bind | Gap | Async I/O accepts cancellation, but CPU-bound tokenization/binding is not cooperatively cancellable. |
| External resource loading | Full offline behavior | Core does not fetch; policy still needs to classify links/file references and keep future adapters offline by default. |
| RTF-to-HTML URL scheme policy | Gap | Parsed RTF links are emitted directly into `href`. |
| HTML private/binary metadata policy | Gap | Round-trip `data-officeimo-*` metadata is not separated from web-safe output. |
| Stable degradation report | Partial | HTML, Markdown, PDF, Reader, and core use different shapes; Word has none. |
| Strict no-loss mode | Gap | No shared `RequireNoLoss()` contract. |
| Operation timeout | Gap | No parser/converter timeout contract. Host cancellation is incomplete. |

## Interoperability corpus

The corpus is under `OfficeIMO.Rtf.Tests/Documents/RtfCorpus`.

| Bucket | Current files | Evidence level |
| --- | ---: | --- |
| Basic text | 1 | Synthetic parse guard. |
| Unicode/code pages | 1 | Windows-1250 synthetic fixture. |
| Formatting | 1 | Synthetic parse guard. |
| Styles/lists | 1 | Synthetic parse guard. |
| Tables | 1 | Synthetic parse guard. |
| Images/objects | 1 | Placeholder fixture, not visual proof. |
| Notes/fields | 1 | Hyperlink field fixture. |
| Sections/layout | 1 | Page setup fixture. |
| Security/pathological | 1 | Unknown destination fixture only. |

Missing producer evidence:

- Microsoft Word versions;
- LibreOffice Writer;
- Outlook HTML encapsulation;
- Google Docs import/export;
- macOS TextEdit/RTFD;
- EHR/CRM/helpdesk generators;
- commercial library output;
- adversarial/fuzz-generated inputs;
- target reopen and visual baselines.

## Evidence rules

- `Full` requires a public API, focused test, and a real producer fixture where interoperability matters.
- A degraded conversion must emit a structured diagnostic or explicit placeholder.
- Lossless syntax preservation does not imply semantic or visual support.
- PDF import is extractive and must not be described as lossless visual reconstruction.
- High-level RTF workflows should route through existing `OfficeIMO.Word` engines and return combined diagnostics.
- The matrix should be generated from or checked against a machine-readable capability manifest so it cannot drift behind published packages again.
