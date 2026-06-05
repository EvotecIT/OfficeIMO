# OfficeIMO Word HTML Roadmap

## Goal

Make `OfficeIMO.Word.Html` the default .NET choice for high-fidelity, safe, and maintainable conversion between HTML and Word documents.

The library should preserve the content people expect from real documents, expose predictable OfficeIMO-owned APIs, and produce evidence through generated DOCX/HTML artifacts rather than only unit-level claims.

## Current Strengths

- Bidirectional conversion exists: HTML to `WordDocument` and `WordDocument` to HTML.
- The HTML import path already handles headings, paragraphs, inline formatting, links, images, SVG, lists, tables, captions, notes, headers, footers, and sections.
- CSS support covers inline declarations, stylesheet content, stylesheet files, remote stylesheets, class-to-style mapping, paragraph spacing, indentation, colors, line height, white space, list styles, table widths, borders, cell spacing, and physical plus logical alignment.
- Resource handling already has caller-provided `HttpClient`, timeout support, data URI images, embedded images, external image links, SVG support, per-image/image-total/CSS-total byte limits, declared image and stylesheet content-type validation, and URI scheme/host policy controls.
- The export path already emits metadata, footnotes, endnotes, optional headers/footers, optional comments, images, SVG, lists, tables, paragraph/run classes, inline run colors/highlights, spacing, indentation, and optional default CSS.
- Existing tests cover many real contracts across HTML import, Word export, style mapping, table behavior, image handling, links, notes, async/cancellation, and whitespace.
- The current supported feature set is documented in `Docs/officeimo.word-html-support-matrix.md`.

## Recently Fixed Issues

- HTML table import now computes grid width with existing row and column spans before allocating the Word table. This fixes valid tables where a `rowspan` reserves a column and later rows add more cells.
- CSS color parsing now accepts modern `rgb()` syntax with space-separated channels, percentage channels, and slash alpha notation. Alpha is ignored because Word run color is opaque.
- HTML import now maps `break-before`, `break-after`, and legacy `page-break-before` / `page-break-after` CSS values to Word page breaks, including container-level `break-after` without duplicating breaks on every child paragraph.
- HTML import now accepts `data-bookmark` as an explicit bookmark source alongside `id` and `name`, so generated documents can preserve stable anchor names without overloading visible element IDs.
- HTML table import now maps CSS auto margins and the table `align` attribute to Word table justification, including common `margin: 0 auto` centering and right alignment through `margin-left:auto`.
- Image import now honors width-only sizing by preserving natural aspect ratio, and maps common percentage image widths against the Word section content width.
- Image import/export now preserves HTML `title` metadata through `WordImage.Title`, including saved DOCX package reload, while continuing to map `alt` through `WordImage.Description`.
- List import now accepts `!important` on `list-style-type`, quoted dash markers from Markdown/editor HTML, and OpenXML-backed international ordered list styles such as Russian, Hebrew, Arabic, Hiragana, and Katakana. List export now preserves those formats when `IncludeListStyles` is enabled.
- Markdown and editor task-list HTML now imports `input[type=checkbox]` list markers as native Word checkbox content controls while preserving the surrounding list item text.
- HTML import now maps text-like `input` elements to native Word structured document tags, preserving the input value plus alias/tag metadata from label-style attributes.
- HTML import now maps single-select `select` elements to native Word dropdown list content controls and `textarea` elements to structured document tags, preserving selected/value text plus alias/tag metadata where available.
- Word-to-HTML export now preserves native Word checkbox content controls as disabled HTML checkbox inputs, including checked state and available alias/tag metadata, instead of leaking checkbox glyphs or dropping the control.
- Word-to-HTML export now preserves native Word structured text controls, dropdown lists, combo boxes, and date pickers as disabled HTML form controls with selected values, option lists, and available alias/tag metadata. Multiline structured text controls now export as `textarea` elements so line breaks round-trip through HTML import.
- HTML import now records conversion diagnostics for skipped or degraded image resources through `HtmlToWordOptions.Diagnostics` and `DiagnosticHandler`, replacing console-only failure reporting while preserving alt-text fallback behavior.
- HTML import now supports `HtmlToWordOptions.MaxImageBytes` for remote, local, data URI, and SVG image resources, producing `ImageResourceTooLarge` diagnostics and preserving alt-text fallback when an image exceeds the configured limit.
- HTML import now validates declared image content types for remote image resources and data URI images through `HtmlToWordOptions.ValidateImageContentTypes` and `AllowedImageContentTypes`, producing `ImageContentTypeRejected` diagnostics and preserving alt-text fallback when an image type is rejected.
- HTML import now supports `HtmlToWordOptions.MaxTotalImageBytes` across a single import operation, producing `ImageResourceBudgetExceeded` diagnostics and preserving alt-text fallback when an image would exceed the remaining total budget.
- HTML import now supports image URI policy controls through `HtmlToWordOptions.AllowedImageUriSchemes` and `AllowedImageHosts`, producing `ImageResourceRejectedByPolicy` diagnostics before disallowed resources are fetched or externally linked.
- HTML import now reports `UnsupportedCssDeclaration` diagnostics for unsupported effective inline and stylesheet CSS properties while avoiding warnings for properties that the importer maps today.
- HTML import now reports `UnsupportedCssValue` diagnostics for supported CSS properties whose effective values cannot be mapped, and `HtmlToWordOptions.UnsupportedCssHandling` lets callers ignore, warn, or fail conversion on unsupported CSS.
- A grounded support matrix now documents current HTML tag, CSS, table, list, image/resource, option, and export coverage.
- HTML import now skips non-rendered `script` and `template` content instead of leaking it into the Word document as visible text, producing `HtmlElementSkipped` diagnostics.
- HTML import now supports structural conversion limits through `MaxHtmlNodes`, `MaxHtmlDepth`, `MaxCssBytes`, and `MaxTableCells`, producing error diagnostics and `HtmlConversionLimitException` before expensive or unsafe conversion work continues.
- Word-to-HTML export now preserves ordinary Word bookmarks as HTML `id` anchors on paragraphs, headings, list items, code blocks, figures, and horizontal rules while keeping structural bookmark wrappers intact.
- Word-to-HTML export can now preserve Word section/page metadata when `WordToHtmlOptions.IncludeSectionMetadata` is enabled, wrapping each section with page size, orientation, margin data attributes, and Word-like CSS dimensions.
- Word-to-HTML export can now preserve custom document properties as typed HTML meta tags when `WordToHtmlOptions.IncludeCustomProperties` is enabled.
- Word-to-HTML table export now preserves leading Word table header rows as `<thead>` / `<th>` markup while keeping ordinary tables in the legacy flat row shape.
- Word-to-HTML table export now emits `scope="col"` on generated header-row `th` cells, giving browser and assistive-technology consumers explicit column-header semantics for Word repeated header rows.
- Word-to-HTML table export now emits adjacent `Caption`-style table captions as semantic `<caption>` elements, and HTML import with `TableCaptionPosition.Below` now places body-level captions after the table instead of inside the last table cell.
- HTML table import now inserts body-level tables directly into the section when no current paragraph exists, avoiding synthetic empty paragraphs before table-only HTML.
- Word-to-HTML table export can now preserve Word table column widths as `<colgroup>` / `<col>` markup when `WordToHtmlOptions.IncludeTableColumnGroups` is enabled, including fractional percentage widths.
- Word-to-HTML export now preserves native Word endnotes as an `endnotes` section with stable `en*` anchors, controlled independently by `WordToHtmlOptions.ExportEndnotes`, and HTML import recreates exported `section.endnotes` references as native Word endnotes.
- Word-to-HTML export now restores imported HTML abbreviations from either footnote or endnote metadata as semantic `abbr title` markup, avoiding visible note-reference leaks in abbreviation round-trips.
- HTML `blockquote cite` import now preserves the source citation URL through immediate Word-to-HTML export and saved DOCX package reload as a `blockquote cite` attribute instead of leaking the generated citation note as a visible footnote reference.
- HTML `time datetime` import now preserves the source machine-readable value through immediate Word-to-HTML export even when the visible label differs from the `datetime` attribute.
- HTML `del` and `ins` import now preserves source deleted/inserted inline semantics through immediate Word-to-HTML export while generic Word strike/underline formatting still exports as visual `s`/`u` tags.
- HTML `mark` import now preserves source highlighted inline semantics through immediate Word-to-HTML export while generic Word highlight formatting remains available as opt-in highlight CSS.
- Word-to-HTML export can now preserve Word comments when `WordToHtmlOptions.ExportComments` is enabled, emitting linked comment references plus a `comments` section with author, initials, date, and reply metadata; HTML import recreates that exported comment section as native Word comments with nested replies.
- Word-to-HTML export can now preserve Word section headers and footers when `WordToHtmlOptions.ExportHeadersAndFooters` is enabled, emitting semantic `header` and `footer` regions with section index and header/footer type metadata; HTML import rehydrates those exported regions back into native section headers and footers.
- Word-to-HTML export can now emit reusable list definition CSS when `WordToHtmlOptions.IncludeListDefinitions` is enabled, assigning stable `word-list-*` classes and Word list-level metadata to generated `ol` and `ul` elements.
- HTML import now maps document-level `html` / `body` language attributes to `WordDocument.Settings.Language`, and Word-to-HTML export emits that setting as the root `html lang` attribute.
- HTML import now maps element-level `lang` / `xml:lang` values to Word run language metadata, including inherited language values from parent containers.
- Word-to-HTML export now emits run-level Word language metadata as `span lang` attributes when the run language differs from the document language, allowing element-level HTML language attributes to round-trip.
- HTML import can now emit opt-in accessibility diagnostics through `HtmlToWordOptions.EnableAccessibilityDiagnostics` for missing image alternate text, weak or empty link text, skipped heading levels, and likely data tables without header cells.
- Core Word table generation now writes table-look conditional formatting through the `w:tblLook` bitmask instead of validator-rejected expanded attributes, keeping generated table DOCX packages OpenXML-valid while preserving the `ConditionalFormatting*` API behavior.
- Word-to-HTML export internals now keep note emission, section metadata, bookmark IDs, checkbox controls, and style-definition CSS in focused partials so the main converter stays below the structure-check threshold.
- HTML import now uses one shared linked-stylesheet loading path for full document, body append, header append, footer append, and body-level `<link rel="stylesheet">` elements, so stylesheet diagnostics, base URL handling, and cancellation behavior stay consistent.
- `OfficeIMO.Reader.Html` now preserves the full nested `HtmlToMarkdownOptions` clone contract when registered or used directly, so conversion controls such as input limits, markdown writer options, transforms, custom element converters, and visual round-trip hints are not silently dropped by the adapter.
- HTML import now supports stylesheet URI policy controls through `HtmlToWordOptions.AllowedStylesheetUriSchemes` and `AllowedStylesheetHosts`, emits deterministic `StylesheetResourceRejectedByPolicy` diagnostics before blocked stylesheets are loaded, and applies `MaxCssBytes` to file and remote stylesheets before buffering the full CSS body when length metadata is available.
- HTML import now supports `HtmlToWordOptions.MaxTotalCssBytes`, `ValidateStylesheetContentTypes`, and `AllowedStylesheetContentTypes`, so linked and configured CSS can be bounded across one conversion and remote stylesheets with disallowed declared media types are skipped with `StylesheetContentTypeRejected` diagnostics.
- HTML import now emits specific linked-stylesheet diagnostics for non-success HTTP statuses, transport failures, and resource-timeout cancellations through `StylesheetHttpStatusRejected`, `StylesheetTransportFailed`, and `StylesheetLoadTimedOut` while preserving true caller cancellation.
- HTML import now keys parsed stylesheet caching by CSS content hash instead of mutable source URL or path, so repeated identical CSS still reuses parsed rules without reusing stale rules when a local or remote stylesheet changes.
- HTML import now emits source-aware linked-stylesheet diagnostics when document-provided links are disabled, when a stylesheet `<link>` has no `href`, and when unsupported stylesheet URI schemes are rejected by policy.
- `HtmlToWordOptions` now provides named import profiles for default OfficeIMO behavior, bounded offline untrusted HTML ingestion, and trusted document stylesheet loading, plus `Clone()` support for reusable option templates without carrying runtime diagnostics between conversions.
- `OfficeIMO.Reader.Html` now provides named adapter profiles for OfficeIMO-default, portable Markdown, and bounded untrusted HTML ingestion, plus public `ReaderHtmlOptions.Clone()` support so registered handler templates and direct-read templates stay independent.
- HTML import now preserves `select multiple` values as structured document tag text, keeping all selected option values and avoiding the browser-inaccurate single-select fallback that picked only one value or defaulted to the first option.
- HTML import now skips unsupported embedded media/widget elements (`iframe`, `object`, `embed`, `video`, `audio`, and `canvas`) with `HtmlEmbeddedContentSkipped` diagnostics instead of leaking fallback text into the generated Word document.
- HTML import now maps named radio groups to a single Word dropdown list content control, preserving the checked value, using labels when radio values are omitted, and avoiding a false first-option selection for unselected groups.
- HTML import now preserves visible value-bearing input types such as `number`, `time`, `datetime-local`, `month`, `week`, `color`, and `range` as structured document tags while continuing to ignore hidden, file, button, submit, reset, and image controls as non-document content.
- HTML import now emits unsupported-value diagnostics for mapped `direction` and `border-collapse` CSS values that cannot be represented in Word output.
- HTML import now maps logical `text-align:start` and `text-align:end` through inherited `dir` / CSS `direction` metadata for paragraphs and table-cell content instead of reporting them as unsupported values.
- HTML import now preserves `meter` and `progress` values as structured document tags, using fallback element text when no explicit value is present.
- HTML import now reports raw HTML comments through `HtmlCommentSkipped` diagnostics by default and can opt in to import non-empty raw comments as native Word comments through `HtmlToWordOptions.ImportHtmlComments`, `HtmlCommentAuthor`, and `HtmlCommentInitials`; native Word comment-section import remains the richer comment round-trip path.
- HTML table import now maps the legacy `cellspacing` attribute and CSS `border-spacing` to Word table cell spacing, with unsupported `border-spacing` values reported through CSS value diagnostics.
- Word-to-HTML table export now emits Word table cell spacing as CSS `border-spacing`, using `border-collapse:separate` when exported table borders are present so spacing survives HTML rendering and HTML import round-trips.
- HTML table import now maps `td`/`th` CSS `vertical-align:top|middle|bottom` and the legacy `valign` attribute to Word table-cell vertical alignment, matching the existing Word-to-HTML export path.
- Word-to-HTML table export now emits a final table row as `<tfoot>` when the Word table has last-row conditional formatting enabled, and HTML import preserves `tfoot` intent through the same Word table-look flag.
- Word-to-HTML structural bookmark export now keeps the original block shape inside semantic wrappers such as `article`, so bookmarked headings export as `<article id="..."><h1>...</h1></article>` instead of flattening heading text into the wrapper.
- The HTML artifact gallery now writes representative source HTML, generated DOCX, and round-trip HTML artifacts while validating the generated DOCX package with OpenXML validation.
- HTML definition lists now round-trip through semantic `dt`/`dd` marker paragraphs, including table-cell cases, so imported descriptions no longer export as blockquotes.
- HTML table-cell block imports now preserve earlier block paragraphs instead of clearing the cell every time another block child is added, and Word-to-HTML table-cell export now avoids duplicate definition-list items when OfficeIMO exposes empty and non-empty wrappers for the same underlying cell paragraph.
- HTML list-item import now preserves source order when block children such as paragraphs and tables appear inside one `li`, so a table following list-item detail text no longer jumps ahead of that detail paragraph or behind later body content.
- HTML list import/export now preserves common editor marker glyphs, including quoted asterisk and plus bullets, as Word list marker text and CSS `list-style-type` values.
- HTML list import/export now preserves arbitrary quoted CSS marker strings as custom Word bullet marker text and quoted `list-style-type` values.
- HTML figure import now normalizes `figcaption` elements after figure media/content, so both leading and trailing captions round-trip through Word's image-plus-caption export pattern as semantic `figure` / `figcaption` HTML.
- HTML blockquote import now styles generated paragraphs from the active conversion scope, so blockquotes inside table cells no longer lose quoted text or export as ordinary paragraphs.
- HTML blockquote citation import now keeps the Word note metadata while also restoring the original `cite` attribute during immediate Word-to-HTML export and after saved DOCX package reload.
- HTML deleted and inserted inline text now round-trips as semantic `del` and `ins` tags instead of degrading to generic strike and underline tags.
- HTML marked inline text now round-trips as semantic `mark` tags instead of degrading to generic highlight styling.
- CSS text-decoration import now maps underline style variants from shorthand and longhand values, clears inherited decoration on `text-decoration:none`, and keeps unsupported decoration color tokens visible through value diagnostics.

## Capability Gaps To Close

- CSS cascade quality: continue improving selector specificity, inherited computed styles, shorthand expansion, `!important`, style precedence, and broader value-level diagnostics for unsupported or degraded declarations.
- CSS resource robustness: continue aligning linked stylesheets with the broader resource pipeline through optional prefetching and broader non-image resource budget coverage beyond CSS.
- Layout and page controls: cover remaining margin and flow placement cases around constrained containers.
- Lists: cover CSS formats that Word does not expose as native numbering, richer nested restart/continuation rules, and robust non-standard list markup produced by editors.
- Tables: continue hardening irregular row/column spans, malformed nested tables in list items, richer collapsed/separate border conflict behavior, multiple-row footer semantics, and table width inside constrained containers.
- Images and resources: add safer prefetching and broader raster/vector formats.
- Semantics and forms: deepen support for `abbr`, `cite`, `dfn`, `q`, `time`, structural elements, bidirectional text, richer specialized form export where practical, native-style multi-select export where practical, and richer reciprocal form behavior; continue expanding richer nested `blockquote` content and richer non-image `figure` content where Word has a durable representation.
- Word-to-HTML fidelity: preserve richer comment range fidelity, round-trip richer header/footer content, style definitions, richer list definitions, table styles, images, richer note formatting, and accessibility metadata more completely.
- Security and reliability: keep parser/resource limits current, prevent regex and CSS parsing DoS paths, and continue expanding warning coverage for skipped or degraded content as new unsupported HTML/resource cases are identified.
- Developer experience: keep the support matrix current, broaden conversion diagnostics, sample gallery, benchmark suite, and round-trip fixture corpus.
- HTML adapter options: keep `OfficeIMO.Reader.Html` pass-through documented as a public adapter contract and continue adding examples for custom transforms, element converters, visual round-trip hints, and Reader chunking profiles.

## Plan

### Phase 1: Baseline And Contracts

- Keep the documented support matrix current for HTML tags, CSS properties, Word features, image formats, and export features.
- Add golden input fixtures from real editors: Markdown-generated HTML, browser copy/paste, Outlook, Word exported HTML, Google Docs, CMS article HTML, code blocks, email snippets, and reporting tables.
- Add artifact validation that opens generated DOCX with OpenXML validation and extracts expected Word structure.
- Add HTML export snapshots for representative Word documents.

### Phase 2: HTML To Word Fidelity

- Build a computed-style layer that merges inline styles, embedded stylesheets, external stylesheets, classes, inheritance, and specificity before conversion.
- Expand CSS property support around fonts, colors, margins, padding, borders, table layout, white space, direction, and text transforms.
- Harden table import with a dedicated table grid model that owns span placement, row groups, column groups, and malformed HTML recovery.
- Expand list import with custom markers, additional numbering formats, robust nested list restarts, and editor-specific list cleanup.
- Add diagnostics for unsupported or degraded HTML, with strict and lenient modes.

### Phase 3: Word To HTML Fidelity

- Export section and page metadata as structured HTML plus CSS where possible.
- Export and round-trip richer header/footer content, richer comment content, bookmarks, richer footnote/endnote content, fields, table styles, richer list definitions, and document properties.
- Generate cleaner, stable HTML with readable class names and optional embedded CSS.
- Continue the HTML accessibility pass with bidirectional text, richer table semantics, and broader export-side accessibility metadata.

### Phase 4: Resource, Security, And Performance

- Add a resource pipeline with configurable allow lists, caching, parallel prefetch, and deterministic error handling.
- Expand conversion limits beyond current DOM size, nesting depth, table size, CSS size, image size, and total resource byte controls as new hostile-input cases are discovered.
- Benchmark large documents, image-heavy documents, deeply nested lists, large tables, and hostile CSS/HTML inputs.
- Keep the reusable conversion logic inside `OfficeIMO.Word.Html` and keep examples, docs, and future PowerShell surfaces thin.

### Phase 5: Market-Ready Polish

- Publish a converter gallery with input HTML, generated DOCX, exported HTML, screenshots, and validation results.
- Ship cookbook examples for web content, reports, emails, legal documents, invoices, knowledge-base pages, and code-heavy documentation.
- Add compatibility notes for Word desktop, Word Online, LibreOffice, Google Docs import/export, and HTML produced by common editors.
- Expose stable diagnostics and feature flags so callers can choose strict archival conversion, forgiving editor conversion, or fast best-effort conversion.

## Success Criteria

- Every supported tag and CSS feature has a contract test or artifact proof.
- Generated DOCX validates with OpenXML validation for the supported scenarios.
- HTML export produces stable, readable, browser-renderable output.
- The converter degrades unsupported content with warnings rather than silent data loss or crashes.
- Large and adversarial inputs have clear limits and performance evidence.
- Public APIs stay OfficeIMO-owned, documented, and reusable without requiring callers to manipulate OpenXML directly.
