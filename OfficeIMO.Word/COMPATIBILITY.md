# OfficeIMO.Word Capability And Compatibility Matrix

This document tracks where `OfficeIMO.Word` is already strong, where support is partial, and where explicit feature work is still needed for mature Word document automation workflows.

It is intentionally honest. "Partial" means usable, not "done".

## Current Matrix

| Area | Status | Notes |
| --- | --- | --- |
| Document create/load/save | Supported | File, stream, byte-array, async, encrypted, macro-enabled, and read-only load/save paths are available. Recent regression coverage protects shared-read async loads, stream-save OpenOffice compatibility, comparer temporary-output ownership, and cloner source-stream disposal. |
| Legacy binary DOC | Partial | Supported Word 97-2003 `.doc` files route through the normal `WordDocument.Load(...)` path and explicit `LoadLegacyDocWithReport(...)` diagnostics are available. The dependency-free importer projects body paragraphs/runs, common direct character and paragraph formatting, built-in paragraph styles, simple tables, single-section page setup, paragraph-boundary multi-section page setup, and OLE document properties into the normal Word model. Native `.doc` saving is available for the supported simple subset: paragraphs, common run/paragraph formatting, tabs, line/page breaks, simple body tables, single-section and paragraph-boundary multi-section page setup, and scalar document properties. Unsupported or preserve-only features such as macros, embedded OLE objects, fast/quick-save state, header/footer stories, footnotes, endnotes, comments, text boxes, images, complex table formatting, merged/nested tables, and richer multi-section shapes are reported or blocked before silent data loss. |
| Paragraphs and runs | Supported | Paragraph creation, run formatting, fonts, colors, highlights, tabs, spacing, borders, shading, line breaks, lists, bookmarks, fields, hyperlinks, and fluent helpers are practical production surfaces. |
| Tables | Supported | Table creation, built-in styles, merge/split, nested tables, widths, row heights, borders, shading, repeat headers, page-break behavior, and object-table helpers are supported. Image export now projects direct, base style-inherited, first-row, banded, and corner table-cell fills and borders, including row/column band sizes, band1/band2 selection, and first/last row/column corner precedence, through shared drawing primitives. Complex externally authored table mutation still needs broader corpus coverage. |
| Sections, headers, footers, and page setup | Supported | Page size, orientation, margins, columns, section handling, headers/footers, odd/even/first variants, page numbers, watermarks, and background color are supported. Mixed imported-document scenarios remain a compatibility proof target. |
| Images and drawing | Partial | Images from file, stream, base64, URL, alt text, sizing, wrapping, crop, transparency, flip/rotate, position, text boxes, and basic shapes are supported. Exact layout fidelity, anchored image interactions, SmartArt, and advanced drawing behaviors remain partial. |
| Fields, TOC, bookmarks, and links | Partial | Common field authoring/read/update/remove, TOC helpers, page fields, merge fields, bookmarks, document variables, bibliography sources, hyperlinks, and cross-references are available. Full field evaluation and complex nested field updates remain partial. |
| Mail merge and templates | Partial | Merge-field support, formatting-preserving simple and complex field replacement, repeated table-row merge regions, grouped table-row regions with separate group/detail templates, repeated block regions for paragraph/table content, nested repeated block regions with per-row child data, body/header/footer/table-cell/block content-control conditional template blocks with nested-region support, template inspection/validation diagnostics for fields, conditionals, and repeated regions, batch output from template files, Custom XML-bound content-control refresh/fill/update workflows, and content-control form-map fill/extraction plus preflight validation across common SDTs including picture controls and repeating-section item text are available. Richer section regions, additional nested region shapes, broader SDT mapping scenarios, and additional formatting-preservation scenarios remain priority work. |
| Notes, comments, and revisions | Partial | Footnotes/endnotes, comments, revision settings, inserted/deleted run helpers, accept/reject APIs, and visible-markup conversion exist. Threaded/resolved comment metadata and full tracked-change workflows remain partial. |
| Content controls and forms | Partial | Common structured document tags are available, including check boxes, date pickers, dropdowns, combo boxes, picture controls, rich text, repeating sections, and tag/alias lookups. Form-map fill/extraction and preflight validation now cover text, checkbox, date picker, dropdown, combobox, picture controls, repeating-section item text, and ambiguous tag/alias key mappings, including picture replacement from files or byte-backed values. Richer binding, advanced validation policies, and advanced mapping workflows need deeper coverage. |
| Charts, shapes, and SmartArt | Partial | Common chart authoring and basic shapes exist; SmartArt is detected and should be treated carefully. Rich chart data updates, chart formatting/readback, SmartArt mutation, grouping, and alignment remain partial. |
| Macros | Partial | VBA projects can be attached, extracted, enumerated, and removed. OfficeIMO does not edit VBA modules or sign macro projects. |
| Digital signatures | Unsupported | Application-level signature metadata can be surfaced, but package signing and signature validation are not implemented. Editing signed documents may invalidate signatures. |
| Protection and security | Partial | Document protection and encrypted OOXML package workflows exist. Broader permission fidelity, signing, and secure automation guidance remain roadmap items. |
| HTML and Markdown conversion | Partial | Adjacent packages provide practical HTML and Markdown workflows. Fidelity corpus coverage and destination-specific diagnostics remain ongoing work. |
| PDF output | Experimental | Word-to-PDF currently depends on temporary PDF infrastructure and should not be treated as a finished compatibility promise yet. |
| Feature inspection | Partial | `InspectFeatures()` reports editable, partially editable, preserved, and unsupported document features, including core content, document variables, bibliography sources, footnotes/endnotes, review metadata, charts, SmartArt diagram package parts, equations, content controls, content-control data bindings, external links, externally linked images, attached template relationships, glossary/building-block metadata, modern comment metadata, web extension/task-pane metadata, ActiveX control package metadata, altChunks, embedded packages, custom XML, VBA projects, and digital signature metadata. Round-trip preservation proof now covers those advanced package signals, including ActiveX XML/binary control metadata; broader corpus-backed preservation proof remains a target. |
| Document compare and diff | Partial | A document comparer exists and lifecycle ownership regressions are covered. `WordDocumentComparer.CompareStructure(...)` now returns deterministic paragraph, table, row, cell, and image findings for review/report automation. Richer run-level formatting diffs, field/content-control diffs, redline reports, and full review workflows remain roadmap items. |

## Current Strengths

- ergonomic code-first `.docx` generation
- first-party dependency-free import of supported legacy `.doc` files into the normal Word model
- broad practical coverage for paragraphs, tables, sections, headers, footers, images, and content controls
- managed Open XML automation without requiring Office installation
- feature inspection for safer automation against unknown documents
- adjacent HTML, Markdown, and PDF packages for destination-specific workflows

## Near-Term Market Readiness Focus

Word-to-PDF fidelity is tracked separately in `OfficeIMO.Word.Pdf` / `OfficeIMO.Pdf`. The non-PDF Word push should focus on the surfaces that can make `OfficeIMO.Word` stand out as a practical open-source Word automation engine:

1. Template and document assembly polish: make the existing merge fields, repeated regions, content-control bindings, validation, and batch output story obvious through scenario docs, proof artifacts, and workflow APIs.
2. Review, redline, and diff workflows: turn comments, revisions, visible markup, and `WordDocumentComparer` into a structured review story with deterministic diff output, review reports, and redline documents.
3. HTML and Markdown conversion fidelity: keep expanding the real-world conversion corpus, diagnostics, support matrix, and generated artifact gallery for Word/HTML/Markdown workflows.
4. Real-document docs and showcase: lead with concrete documents, source inputs, validation status, and known limitations instead of only listing object-model features.

Keep this section aligned with the generated proof-gallery example and the public Word market-readiness docs as implementation slices land.

## Highest-Priority Gaps

1. Build the full template/mail-merge engine: table-row regions, grouped table-row regions with separate group/detail templates, repeated block regions for paragraph/table content, nested repeated block regions with per-row child data, body/header/footer/table-cell/block content-control conditional blocks, template validation for fields, conditionals, and repeated regions, batch output, bound-content-control preflight, Custom XML-bound content-control fill/update, content-control form-map fill/extraction plus preflight validation across common SDTs including picture controls, repeating-section item text, and ambiguous tag/alias key mappings, and formatting-preserving simple/complex merge-field replacement now exist, while richer section regions, additional nested region shapes, broader SDT mapping scenarios, and deeper formatting-preservation coverage remain roadmap items.
2. Expand review workflows: richer revision readback, comment reply/resolution metadata, structured diff output, and redline reports.
3. Grow the real-world document corpus for feature inspection and preservation-sensitive round trips; current inspection now includes editable document variables/bibliography sources, partially editable externally linked image signals, plus preserve-only attached template, glossary/building-block, modern comment, web extension/task-pane, and ActiveX control package signals.
4. Publish lifecycle and compatibility guidance for file, stream, memory, encrypted, macro-enabled, legacy binary, read-only, and unknown-document workflows.
5. Keep PDF/layout validation out of release promises until the PDF stack is ready for that responsibility.

## Suggested Release-Prep Checks

1. Build `OfficeIMO.Word` on all supported target frameworks.
2. Run Word lifecycle, feature-inspection, comments/revisions, content-controls, macros, and document-compare test slices before release candidates.
3. Run Open XML validation on representative generated documents and selected externally authored corpus files.
4. Update this matrix whenever a major Word feature is added or a compatibility gap is closed.
