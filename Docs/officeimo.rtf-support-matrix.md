# OfficeIMO RTF Support Matrix

This matrix is the working product contract for RTF read/write and conversion. It separates what OfficeIMO preserves losslessly, what it binds semantically, and what each adapter can faithfully project today.

Status legend:

- `Full`: covered by public API and focused tests.
- `Partial`: supported with known loss or limited projection.
- `Extractive`: content is recovered for reading/search/AI use, not round-trip fidelity.
- `Planned`: in scope for the RTF market roadmap but not implemented yet.
- `Deferred`: intentionally postponed to avoid colliding with active PDF work.
- `No`: not currently supported.

Conversion class legend:

- `Lossless`: preserves original RTF syntax/tree for untouched content.
- `Semantic`: preserves editable document meaning through `RtfDocument`.
- `Visual`: targets rendered appearance, may lose source structure.
- `Extractive`: recovers useful text/metadata/provenance from a richer source.
- `Diagnostic`: unsupported or degraded content is reported instead of silently disappearing.

## Package Ownership

| Area | Owning package | Contract |
| --- | --- | --- |
| RTF syntax, parser, diagnostics, model, writer, lossless editor | `OfficeIMO.Rtf` | Core engine. No Word/HTML/Markdown/PDF-specific mapping logic belongs here. |
| RTF <=> Word | `OfficeIMO.Word.Rtf` | Thin adapter between `WordDocument` and `RtfDocument`. |
| RTF <=> HTML | `OfficeIMO.Html` | Thin adapter between HTML DOM/CSS policy and `RtfDocument`. |
| RTF <=> Markdown | `OfficeIMO.Rtf.Markdown` | Thin adapter between `RtfDocument` and `MarkdownDoc`. |
| RTF => Reader chunks | `OfficeIMO.Reader.Rtf` | Extraction/chunking adapter for ingestion workflows. |
| RTF <=> PDF | `OfficeIMO.Rtf.Pdf` | Deferred implementation surface while PDF work is active elsewhere. |

## Core RTF Engine

| Capability | Syntax tree | Semantic model | Writer | Lossless edit | Diagnostics | Notes |
| --- | --- | --- | --- | --- | --- | --- |
| Control words, symbols, groups, text, binary payloads | Full | Full | Full | Full | Full | Parser keeps source tokens for lossless output. |
| Unknown destinations | Full | Partial | Full | Full | Full | Unknown syntax is preserved; semantic binding warns when requested. |
| Fonts and colors | Full | Full | Full | Full | Full | Includes core font/color tables and lossless table edits. |
| ANSI code pages and Unicode escapes | Full | Partial | Full | Full | Full | Single-byte Windows ANSI code pages 874 and 1250-1258 are supported. |
| Paragraphs and runs | Full | Full | Full | Partial | Full | Lossless structural editing still needs more operations. |
| Character formatting | Full | Full | Full | Partial | Full | Bold, italic, underline, strike, caps, hidden, offset, spacing, colors, borders. |
| Paragraph layout | Full | Full | Full | Partial | Full | Alignment, indents, spacing, line spacing, pagination, direction, frame metadata. |
| Styles | Full | Partial | Partial | Partial | Full | Rich style fidelity is an active hardening area for Word/HTML/Markdown. |
| Lists and list tables | Full | Partial | Partial | Partial | Full | Basic list semantics exist; cross-format numbering fidelity needs scorecard cases. |
| Tables | Full | Full | Full | Partial | Full | Rows, cells, merges, widths, padding, shading, borders, text flow. |
| Images | Full | Full | Full | Planned | Full | PNG/JPEG data is modeled; HTML export now reports skipped images when embedding is disabled, data is missing, or the format is unsupported. |
| Fields and hyperlinks | Full | Partial | Partial | Planned | Full | Visible field result is usable; instruction fidelity needs adapter-specific tests. |
| Notes, annotations, generated references | Full | Partial | Partial | Planned | Full | Footnote/endnote bodies exist; Markdown projection needs dedicated contract. |
| Headers and footers | Full | Partial | Partial | Planned | Full | Important for Word/HTML fidelity work. |
| Sections and page setup | Full | Full | Full | Partial | Full | PDF implementation changes are deferred. |
| Revisions/comments | Full | Partial | Partial | Planned | Full | Markdown/HTML should expose loss/degradation diagnostics. |
| Shapes and objects | Full | Partial | Partial | Planned | Full | Preserve source, expose placeholders/diagnostics where target format cannot represent them. |

## Adapter Matrix

| Feature | Word | HTML | Markdown | Reader | PDF |
| --- | --- | --- | --- | --- | --- |
| Plain paragraphs/runs | Full | Full | Full | Full | Deferred |
| Bold/italic/strike/underline | Full | Full | Full | Full | Deferred |
| Font family/size/color/highlight | Partial | Partial | Partial | Extractive | Deferred |
| Hyperlinks | Partial | Full | Full | Full | Deferred |
| Bookmarks | Partial | Partial | Planned | Extractive | Deferred |
| Headings/styles | Partial | Partial | Partial | Full | Deferred |
| Ordered/unordered lists | Partial | Partial | Full | Full | Deferred |
| Nested lists and numbering restarts | Partial | Partial | Partial | Partial | Deferred |
| Tables | Full | Full | Full | Full | Deferred |
| Merged cells | Full | Partial | Partial | Partial | Deferred |
| Images | Full | Partial + Diagnostic | Partial | Extractive | Deferred |
| Footnotes/endnotes | Partial | Partial | Planned | Full | Deferred |
| Comments/annotations | Partial | Partial | Planned | Full | Deferred |
| Headers/footers | Partial | Partial | Planned | Full | Deferred |
| Sections/page setup | Full | Partial | Diagnostic | Extractive | Deferred |
| Revisions/tracked changes | Partial | Partial | Planned | Extractive | Deferred |
| Shapes/text boxes/objects | Partial | Partial | Diagnostic | Extractive | Deferred |
| Unknown destinations | Diagnostic | Diagnostic | Diagnostic | Diagnostic | Deferred |

## Golden Corpus Buckets

The corpus lives under `OfficeIMO.Tests/Documents/RtfCorpus`. Every fixture should state its expected conversion class and the target adapters it protects.

| Bucket | Purpose | Initial status |
| --- | --- | --- |
| `basic-text` | Plain paragraphs, run boundaries, line/page breaks, tabs. | Added |
| `unicode-codepages` | Unicode escapes, ANSI code pages, fallback behavior. | Added |
| `formatting` | Character and paragraph formatting. | Added |
| `styles-lists` | Stylesheets, list tables, legacy numbering. | Added |
| `tables` | Table structure, merged cells, borders, shading. | Added |
| `images-objects` | Pictures, OLE/object placeholders, shapes. | Added |
| `notes-fields` | Footnotes, endnotes, annotations, fields, hyperlinks. | Added |
| `sections-layout` | Page setup, columns, headers/footers, section breaks. | Added |
| `security-pathological` | Malformed groups, depth limits, binary payload limits, unknown destinations. | Added |

## Scorecard Rules

- A feature is `Full` only when there is a public API, a focused test, and at least one corpus fixture where applicable.
- A degraded conversion must surface a diagnostic, warning, placeholder, or documented extraction boundary.
- RTF <=> Markdown should prefer a direct `RtfDocument` <=> `MarkdownDoc` bridge over routing through Word or HTML.
- PDF entries remain `Deferred` until the active PDF implementation stream is ready to integrate.
