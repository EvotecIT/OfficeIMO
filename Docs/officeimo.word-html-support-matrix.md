# OfficeIMO Word HTML Support Matrix

This matrix documents the current `OfficeIMO.Word.Html` conversion surface. It is grounded in the converter implementation and the HTML-focused tests under `OfficeIMO.Tests`.

## Status Key

- Supported: Converted intentionally with current behavior covered by code paths and focused tests.
- Partial: Converted for common cases, with known limits or lossy mapping to Word/HTML.
- Planned: Not yet a dependable contract, or only falls through as plain child content.
- Not supported: Ignored for safety or outside the current Word conversion scope.

## HTML To Word Import

| Area | Status | Current behavior |
| --- | --- | --- |
| Document root | Supported | `html` and `body` are parsed through AngleSharp; document-level `lang` / `xml:lang` maps to `WordDocument.Settings.Language`; body-level inline and stylesheet CSS can flow into descendants. |
| Basic blocks | Supported | `p`, `div`, `br`, and `hr` map to Word paragraphs, runs, line breaks, and horizontal-rule style paragraphs. |
| Headings | Supported | `h1` through `h6` map to Word heading styles, with optional heading numbering. |
| Structural sections | Supported | `section`, `article`, `aside`, `nav`, `header`, `footer`, and `main` preserve content; structural IDs can round-trip through bookmarks. |
| Definition lists | Supported | `dl`, `dt`, and `dd` import as paragraphs, with `dd` indented. |
| Block quotes | Supported | `blockquote` imports with quote styling/indentation and can preserve `cite` through note references. |
| Inline formatting | Supported | `strong`, `b`, `em`, `i`, `u`, `s`, `del`, `ins`, `sup`, `sub`, `small`, `big`, `font`, and `span` map to run formatting where Word has an equivalent. |
| Semantic inline text | Supported | `q`, `cite`, `dfn`, `time`, `abbr`, `acronym`, and `mark` preserve recognizable Word character styles or notes where available. |
| Code and keyboard text | Supported | `pre` imports as block code; inline `code`, `kbd`, `samp`, and `tt` import as monospace runs. |
| Ruby text | Partial | `ruby`, `rb`, `rt`, and `rp` keep child text and inline styles, but do not create Word ruby annotation structures yet. |
| Links and anchors | Supported | `a href`, `id`, `name`, and `data-bookmark` map to hyperlinks and bookmarks. |
| Images | Supported | `img` imports from data URIs, local paths, file URLs, remote URLs, and external image links depending on options. |
| SVG | Supported | Inline SVG and SVG image sources can be embedded, with alt-text fallback and diagnostics when skipped. |
| Lists | Supported | `ul`, `ol`, and `li` map to Word lists, including nested lists, `start`, `value`, reversed lists, common `type` values, and CSS `list-style-type`. |
| Task lists | Supported | Markdown/editor `input[type=checkbox]` list markers import as native Word checkbox content controls. |
| Tables | Supported | `table`, `tr`, `th`, `td`, `thead`, `tbody`, `tfoot`, `caption`, `colgroup`, and `col` import with spans, section rows, captions, widths, borders, and cell styling. |
| Figures | Supported | `figure` and `figcaption` import as figure content plus Word caption paragraphs. |
| Stylesheets | Supported | Inline `style`, embedded `<style>`, configured stylesheet contents, configured stylesheet paths, file links, and remote stylesheet links are parsed. |
| Language metadata | Supported | Document-level `html` / `body` language attributes map to the Word document language setting. Element-level `lang` / `xml:lang` values inherit to imported text runs as Word run language metadata. |
| Accessibility diagnostics | Supported | When `HtmlToWordOptions.EnableAccessibilityDiagnostics` is enabled, import emits advisory diagnostics for missing image `alt`, weak or empty link text, skipped heading levels, and likely data tables without `th`, `thead`, or `scope`. |
| Active and inert content | Not supported | `script` and `template` content is skipped with `HtmlElementSkipped` diagnostics; executable behavior is never converted. Unsupported harmless elements fall through to child content unless explicitly handled. |
| Embedded media/widgets | Planned | `iframe`, `object`, `embed`, `video`, `audio`, and `canvas` do not have a Word conversion contract yet. |
| Form controls | Partial | Checkbox task-list markers import as checkbox content controls. Text-like `input` and `textarea` elements import as structured document tags with value and alias/tag metadata. Single-select `select` elements import as dropdown list content controls. Richer input types and multi-select controls remain planned. |

## CSS Import

| Area | Status | Current behavior |
| --- | --- | --- |
| Cascade sources | Supported | Inline styles, embedded stylesheets, configured stylesheets, linked stylesheets, and body/ancestor inherited styles are applied. |
| Selectors | Partial | Common selectors, classes, specificity, inheritance, and `!important` are handled for supported properties; full browser selector/layout behavior is not the target yet. |
| Fonts | Supported | `font`, `font-family`, `font-size`, `font-style`, `font-variant`, `font-weight`, and legacy `font` attributes map to Word run formatting. |
| Colors | Supported | Hex, named colors, `rgb()`, modern space-separated `rgb()`, percentage channels, slash alpha syntax, and background colors map to Word colors where possible. |
| Text decoration | Supported | Bold, italic, underline, strike, highlight, superscript, subscript, small caps, uppercase transforms, and code font mapping are supported. |
| Paragraph alignment | Supported | `text-align`, direction/RTL, indentation, text indent, margins, line height, paragraph spacing, and white-space behavior map to Word paragraph formatting. |
| Page breaks | Supported | `break-before`, `break-after`, and legacy page-break declarations map to Word page breaks. |
| Tables | Supported | Table/cell width, borders, background color, alignment, padding, margins, `border-collapse`, captions, and colgroup widths are handled for common document tables. |
| Lists | Supported | `list-style-type` maps to native Word list formats where OpenXML has a matching numbering style. |
| Images | Supported | Width, height, percentage width, natural aspect ratio, and float/alignment hints are handled for common image placement. |
| Browser layout | Planned | `display`, flexbox, grid, positioning, z-index, floats as layout, media queries, responsive breakpoints, and full box layout are not yet Word layout contracts. |
| Unsupported CSS diagnostics | Partial | Unsupported effective inline and stylesheet CSS properties emit `UnsupportedCssDeclaration` diagnostics. Unsupported values for mapped CSS properties emit `UnsupportedCssValue` diagnostics. `HtmlToWordOptions.UnsupportedCssHandling` can ignore unsupported CSS, warn by default, or stop conversion with `HtmlUnsupportedCssException`. Broader value diagnostics for additional mapped properties remain planned. |

## Image And Resource Import

| Feature | Status | Current behavior |
| --- | --- | --- |
| Embedded images | Supported | Image bytes can be embedded into the DOCX package. |
| External image links | Supported | `ImageProcessingMode.LinkExternal` can preserve external references when the source and dimensions are usable. |
| Data URI images | Supported | Base64 data URI images are supported; SVG text data URIs are supported. |
| Local and file URL images | Supported | Local paths and `file:` URLs resolve directly or relative to `HtmlToWordOptions.BasePath`. |
| Remote images | Supported | HTTP/HTTPS images load through the configured or default HTTP pipeline with optional timeout. |
| SVG | Supported | Inline SVG, SVG files, remote SVG, and SVG data URIs can be embedded. |
| Duplicate resources | Supported | Duplicate image sources are cached inside a conversion operation. |
| Per-image byte limit | Supported | `MaxImageBytes` rejects oversized image resources with `ImageResourceTooLarge`. |
| Aggregate byte budget | Supported | `MaxTotalImageBytes` rejects image resources that would exceed the conversion budget with `ImageResourceBudgetExceeded`. |
| Content-type validation | Supported | `ValidateImageContentTypes` and `AllowedImageContentTypes` reject disallowed declared image types. |
| URI policy | Supported | `AllowedImageUriSchemes` and `AllowedImageHosts` reject disallowed schemes/hosts before fetching or linking. |
| Diagnostics | Supported | Skipped or degraded image resources, unsupported CSS declarations, unsupported mapped CSS values, and opt-in accessibility warnings populate `Diagnostics` and invoke `DiagnosticHandler`. |
| Structural limits | Supported | `MaxHtmlNodes`, `MaxHtmlDepth`, `MaxCssBytes`, and `MaxTableCells` stop conversion with `HtmlConversionLimitException` and error diagnostics. |
| Non-image resource budgets | Planned | CSS and other external resources do not yet share the image byte-budget pipeline. |
| Parallel prefetch | Planned | Resource loading is deterministic today; parallel prefetching remains a performance roadmap item. |

## Table Conversion

| Feature | Import | Export | Notes |
| --- | --- | --- | --- |
| Basic tables | Supported | Supported | Rows and cells convert in both directions. |
| Header/body/footer row groups | Supported | Partial | Import preserves row-group intent where Word can express it; export emits leading Word header rows as `<thead>` / `<th>` and body rows as `<tbody>` for those tables. Footer row-group export remains planned. |
| Captions | Supported | Partial | Import supports caption placement; export emits figure/table-related captions in common cases. |
| Column groups | Supported | Partial | Import uses `colgroup`/`col` widths. Export emits Word column widths as `colgroup`/`col` markup when `IncludeTableColumnGroups` is enabled. |
| Colspan/rowspan | Supported | Supported | Grid spans and vertical merges convert in common and irregular table layouts. |
| Widths and borders | Supported | Supported | Common CSS width and border values are mapped; full browser border conflict resolution is planned. |
| Cell styles | Supported | Supported | Background color, alignment, width, and border styles are handled for common tables. |
| Nested tables | Partial | Partial | Supported when represented by OfficeIMO section/table traversal; deeper malformed editor output needs more fixtures. |

## List Conversion

| Feature | Import | Export | Notes |
| --- | --- | --- | --- |
| Ordered and unordered lists | Supported | Supported | Native Word numbering is used where possible. |
| Nested lists | Supported | Supported | Nested levels are tracked through Word list levels. |
| Starts and continuation | Supported | Supported | `start`, `value`, reversed lists, and continued numbering are handled for common cases. |
| CSS list styles | Supported | Supported | Common bullet and ordered styles map to native OpenXML formats. |
| International ordered styles | Supported | Supported | Russian, Hebrew, Arabic, Hiragana, Katakana, and related OpenXML-backed formats are mapped when available. |
| Custom marker glyphs | Partial | Planned | A few common editor markers such as dash bullets are mapped; arbitrary generated marker content is planned. |
| Task list checkboxes | Supported | Supported | Checkbox task-list HTML imports as Word checkbox controls and exports back as disabled HTML checkboxes. Other native Word form controls export through the Word-to-HTML form-control path. |

## Word To HTML Export

| Area | Status | Current behavior |
| --- | --- | --- |
| Document shell | Supported | Emits HTML, root `lang` when the Word document language is set, head metadata, title, optional default CSS, additional meta tags, additional link tags, and optional body font family. |
| Paragraphs and headings | Supported | Emits `p`, `h1` through `h6`, `blockquote`, `hr`, paragraph classes, alignment, spacing, indentation, background, and border styles. |
| Runs | Supported | Emits `strong`, `em`, `u`, `s`, `sup`, `sub`, `q`, `cite`, `dfn`, `time`, `code`, font spans, run classes, colors, and highlights. |
| Links and bookmarks | Supported | Exports external and anchor hyperlinks as `a href`; ordinary Word bookmarks export as HTML `id` anchors and structural bookmarks export as matching structural elements. |
| Images and SVG | Supported | Emits `img` with base64 data URIs or file/external references; SVG can be emitted inline or as image references depending on options. |
| Lists | Supported | Emits `ol`, `ul`, `li`, `start`, `type`, optional CSS list styles, and optional reusable `word-list-*` definition classes plus a head stylesheet. |
| Tables | Supported | Emits tables with spans, widths, borders, background, alignment, cell styles, optional `colgroup` column widths, and leading header rows as `<thead>` / `<th>` where Word marks them as repeating header rows. |
| Figures and captions | Partial | Exports common image-plus-caption patterns as `figure` and `figcaption`. |
| Notes | Supported | Emits footnote and endnote sections when enabled. Exported `section.footnotes` and `section.endnotes` links import back as native Word footnotes and endnotes. |
| Comments | Partial | When `ExportComments` is enabled, emits linked Word comment references and a `section.comments` list with author, initials, date, and nested reply metadata. Exported comment sections import back as native Word comments with nested replies; arbitrary HTML comment import and richer range fidelity remain planned. |
| Headers and footers | Partial | When `ExportHeadersAndFooters` is enabled, emits non-empty section headers and footers as semantic `header`/`footer` regions with `word-header`/`word-footer` classes, section indexes, and default/first/even type metadata. Exported regions import back into native section headers/footers; richer header/footer content remains planned. |
| Form/content controls | Supported | Emits native Word checkbox, structured text, dropdown list, combo box, and date picker content controls as disabled HTML form controls with selected values, option lists, and available metadata. |
| Sections/page setup | Partial | When `IncludeSectionMetadata` is enabled, export wraps each Word section in a `section.word-section` element with page size, orientation, margin data attributes, and Word-like CSS dimensions. |
| Properties and definitions | Partial | Built-in document properties export as standard metadata. Custom document properties export as typed meta tags when `IncludeCustomProperties` is enabled. Document language exports as root `html lang`. Paragraph/run style definitions export when class output is enabled. List definitions export as reusable CSS when `IncludeListDefinitions` is enabled. Broader run-level language and accessibility metadata export remains planned. |

## Public Options

| Option surface | Status | Key controls |
| --- | --- | --- |
| `HtmlToWordOptions` | Supported | Font family, quote text, default page size/orientation, class styles, list styles, numbering continuation, heading numbering, base path, notes, image processing mode, `HttpClient`, resource timeout, image limits, HTML/CSS/table limits, content-type validation, URI policy, diagnostics, opt-in accessibility diagnostics, unsupported CSS handling, stylesheet paths/contents, pre rendering mode, caption placement, and section handling. |
| `WordToHtmlOptions` | Supported | Font family, font styles, list styles, list definitions, paragraph/run classes, run color/highlight styles, paragraph spacing/indentation styles, footnote and endnote export, header/footer export, comment export, custom property metadata export, section metadata export, table column-group export, image embedding mode, additional meta/link tags, and default CSS. |

## Current Validation Evidence

- HTML-focused tests under `OfficeIMO.Tests` cover import, export, tables, lists, images, links, inline styles, page settings, code blocks, task lists, async/cancellation, and options.
- The current roadmap worktree last ran the broad HTML-focused test suite with `508/508` passing on `net8.0`, `net10.0`, and `net472`.
- The current roadmap worktree also ran the full `OfficeIMO.Tests` project with `8,319` passed / `7` skipped on `net8.0`, `8,319` passed / `7` skipped on `net10.0`, and `8,315` passed / `7` skipped on `net472`.
- This matrix is a contract inventory. Rows marked Supported should continue gaining direct fixture or artifact evidence as the roadmap moves through the remaining phases.
