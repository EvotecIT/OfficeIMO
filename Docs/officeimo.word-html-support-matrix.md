# OfficeIMO Word HTML Support Matrix

This matrix documents the current `OfficeIMO.Word.Html` conversion surface. It is grounded in the converter implementation and the HTML-focused tests under `OfficeIMO.Html.Tests`, `OfficeIMO.Word.Tests`, and `OfficeIMO.Reader.Tests`.

## Status Key

- Supported: Converted intentionally with current behavior covered by code paths and focused tests.
- Partial: Converted for common cases, with known limits or lossy mapping to Word/HTML.
- Not supported: Ignored for safety or outside the current Word conversion scope.

## HTML To Word Import

| Area | Status | Current behavior |
| --- | --- | --- |
| Document root | Supported | `html` and `body` are parsed through AngleSharp; document-level `lang` / `xml:lang` maps to `WordDocument.Settings.Language`; body-level inline and stylesheet CSS can flow into descendants. |
| Basic blocks | Supported | `p`, `div`, `br`, and `hr` map to Word paragraphs, runs, line breaks, and horizontal-rule style paragraphs. Block background, border, margin, and padding map to native paragraph shading, borders, spacing, and indentation; multi-paragraph containers keep one visual frame. |
| Headings | Supported | `h1` through `h6` map to Word heading styles, with optional heading numbering. |
| Structural sections | Supported | `section`, `article`, `aside`, `nav`, `header`, `footer`, and `main` preserve content; structural IDs can round-trip through bookmarks. |
| Definition lists | Supported | `dl`, `dt`, and `dd` import as semantic term/description paragraphs, with `dd` indented and marker styles preserved for Word-to-HTML round-trip export. |
| Block quotes | Supported | `blockquote` imports with quote styling/indentation in body and table-cell scopes, preserves `cite` through note references, restores imported citation URLs as `blockquote cite` attributes during immediate Word-to-HTML export and saved DOCX package reload, and exports scoped quoted paragraphs back as semantic blockquotes. |
| Inline formatting | Supported | `strong`, `b`, `em`, `i`, `u`, `s`, `del`, `ins`, `sup`, `sub`, `small`, `big`, `font`, and `span` map to run formatting where Word has an equivalent. Imported `del` and `ins` runs preserve their source semantic tags during immediate Word-to-HTML export, while generic Word strike/underline formatting exports as visual `s`/`u` tags. |
| Semantic inline text | Supported | `q`, `cite`, `dfn`, `time`, `abbr`, `acronym`, and `mark` preserve recognizable Word character styles or notes where available. Imported `mark` tags round-trip through immediate Word-to-HTML export, imported `time datetime` values round-trip through immediate Word-to-HTML export, and abbreviation title metadata can round-trip through footnote or endnote-backed import. |
| Code and keyboard text | Supported | `pre` imports as block code; inline `code`, `kbd`, `samp`, and `tt` import as monospace runs. |
| Ruby text | Supported | `ruby` elements import as native Word ruby annotation structures and export as `ruby`/`rb`/`rt`. Multiple explicit `rb`/`rt` pairs remain distinct and ordered. Unequal or ambiguous explicit segment counts are combined only with `HtmlRubyPairingApproximation`; ruby without an annotation remains visible text. Rich formatting inside one ruby segment is not a stable contract. |
| Links and anchors | Supported | `a href`, `id`, `name`, and `data-bookmark` map to hyperlinks and bookmarks. |
| Images | Supported | `img` imports from data URIs, local paths, file URLs, remote URLs, and external image links depending on options. `alt` maps to `WordImage.Description`, and `title` maps to durable `WordImage.Title` metadata, including duplicate-source images that share the same binary part. |
| SVG | Supported | Inline SVG and SVG image sources can be embedded, with alt-text fallback and diagnostics when skipped. |
| Lists | Supported | `ul`, `ol`, and `li` map to Word lists, including nested lists, `start`, `value`, reversed lists, common `type` values, CSS `list-style-type`, and source-ordered block children such as paragraphs and tables inside list items. |
| Task lists | Supported | Markdown/editor `input[type=checkbox]` list markers import as native Word checkbox content controls. |
| Tables | Supported | `table`, `tr`, `th`, `td`, `thead`, `tbody`, `tfoot`, `caption`, `colgroup`, and `col` import with spans, section rows, captions, widths, borders, cell spacing, and cell styling. Body-level tables import directly without synthetic placeholder paragraphs. |
| Figures | Supported | `figure` and `figcaption` import as figure content plus Word caption paragraphs. Leading and trailing figure captions normalize to Word's image-plus-caption order so common image figures export back as semantic `figure` / `figcaption` HTML. |
| Stylesheets | Supported | Inline `style`, embedded `<style>`, configured stylesheet contents, configured stylesheet paths, head and body stylesheet links, file links, and remote stylesheet links are parsed. Document-provided links remain opt-in through `AllowDocumentStylesheetLinks`. Stylesheet URI schemes, hosts, declared content types, per-stylesheet byte limits, and aggregate CSS bytes can be restricted before or during external CSS loading. |
| Language metadata | Supported | Document-level `html` / `body` language attributes map to the Word document language setting. Element-level `lang` / `xml:lang` values inherit to imported text runs as Word run language metadata. |
| Accessibility diagnostics | Supported | When `HtmlToWordOptions.EnableAccessibilityDiagnostics` is enabled, import emits advisory diagnostics for missing image `alt`, weak or empty link text, skipped heading levels, and likely data tables without `th`, `thead`, or `scope`. |
| Active and inert content | Not supported | `script` and `template` content is skipped with `HtmlElementSkipped` diagnostics; raw HTML comments are skipped with `HtmlCommentSkipped` diagnostics unless `ImportHtmlComments` is enabled. Executable behavior is never converted. Unsupported harmless elements fall through to child content unless explicitly handled. |
| Embedded media/widgets | Not supported | `iframe`, `object`, `embed`, `video`, `audio`, and `canvas` are skipped with `HtmlEmbeddedContentSkipped` diagnostics instead of leaking fallback text into the Word document. |
| Form controls | Supported | Checkbox task-list markers import as checkbox content controls. Text-like and visible value-bearing `input` elements (`text`, `search`, `email`, `url`, `tel`, `password`, `number`, `time`, `datetime-local`, `month`, `week`, `color`, and `range`), `textarea`, `meter`, and `progress` elements import as structured document tags with value and alias/tag metadata. Date inputs import as date picker content controls. Datalist-backed inputs import as combo boxes. Single-select `select` elements and named radio groups import as dropdown list content controls. Multi-select values remain newline-separated text. Hidden, file, button, submit, reset, and image controls are deliberately not document content. Word export emits the supported native controls as disabled HTML controls; browser submission behavior is not a Word contract. |

## CSS Import

| Area | Status | Current behavior |
| --- | --- | --- |
| Cascade sources | Supported | Inline styles, embedded stylesheets, configured stylesheets, linked stylesheets, and body/ancestor inherited styles are applied. |
| Selectors | Partial | Common selectors, classes, specificity, inheritance, and `!important` are handled for supported properties; full browser selector/layout behavior is not the target yet. |
| Fonts | Supported | `font`, `font-family`, `font-size`, `font-style`, `font-variant`, `font-weight`, and legacy `font` attributes map to Word run formatting. |
| Colors | Supported | Hex, named colors, `rgb()`, modern space-separated `rgb()`, percentage channels, and slash alpha syntax map to Word colors. Inline text backgrounds use exact run shading by default; `HtmlTextBackgroundMode.NearestHighlight` provides an explicit palette approximation mode with diagnostics. |
| Text decoration | Supported | Bold, italic, underline, CSS underline style variants, strike, highlight, superscript, subscript, small caps, uppercase transforms, and code font mapping are supported. `text-decoration:none` clears inherited underline/strike formatting. |
| Paragraph alignment | Supported | Physical `text-align` values plus logical `start`/`end` values, direction/RTL, indentation, text indent, physical and logical margin/padding properties, line height, paragraph spacing, and white-space behavior map to Word paragraph formatting. Logical inline start/end values resolve against inherited direction. |
| Page breaks | Supported | `break-before`, `break-after`, and legacy page-break declarations map to Word page breaks. |
| Tables | Supported | Table/cell width, borders, background color, horizontal and table-cell vertical alignment, padding, margins, `border-collapse`, `border-spacing`, captions, and colgroup widths are handled for common document tables. |
| Lists | Supported | `list-style-type` maps to native Word list formats where OpenXML has a matching numbering style. |
| Images | Supported | Width, height, percentage width, natural aspect ratio, and float/alignment hints are handled for common image placement. |
| Browser layout | Not supported | `display`, flexbox, grid, positioning, z-index, floats as layout, media queries, responsive breakpoints, and full box layout are not Word layout contracts. |
| Unsupported CSS diagnostics | Supported | Unsupported effective inline and stylesheet CSS properties emit `UnsupportedCssDeclaration`. A property that is mapped only for another element context, such as table border spacing on a span, is also reported rather than treated as supported. Unsupported values for mapped properties emit `UnsupportedCssValue`, including length, color, border, font, decoration, alignment, list, break, float, and table-spacing profiles. `HtmlToWordOptions.UnsupportedCssHandling` can ignore, warn, or stop conversion with `HtmlUnsupportedCssException`. |

## Image And Resource Import

| Feature | Status | Current behavior |
| --- | --- | --- |
| Embedded images | Supported | Image bytes can be embedded into the DOCX package. |
| External image links | Supported | `ImageProcessingMode.LinkExternal` can preserve external references when the source and dimensions are usable. |
| Data URI images | Supported | Base64 data URI images are supported; SVG text data URIs are supported. |
| Local and file URL images | Supported | Local paths and `file:` URLs resolve directly or relative to `HtmlToWordOptions.BasePath`. |
| Remote images | Supported | HTTP/HTTPS images load through the configured or default HTTP pipeline with optional timeout and bounded concurrency through `MaxConcurrentResourceLoads`. Imports with an aggregate image budget load sequentially to preserve deterministic pre-read budget checks. |
| SVG | Supported | Inline SVG, SVG files, remote SVG, and SVG data URIs can be embedded. |
| Duplicate resources | Supported | Duplicate image sources are cached inside a conversion operation while preserving per-image `alt` and `title` metadata. |
| Per-image byte limit | Supported | `MaxImageBytes` rejects oversized image resources with `ImageResourceTooLarge`. |
| Aggregate byte budget | Supported | `MaxTotalImageBytes` rejects image resources that would exceed the conversion budget with `ImageResourceBudgetExceeded`. |
| Content-type validation | Supported | `ValidateImageContentTypes` and `AllowedImageContentTypes` reject disallowed declared image types. |
| URI policy | Supported | `AllowedImageUriSchemes` and `AllowedImageHosts` reject disallowed schemes/hosts before fetching or linking. |
| Diagnostics | Supported | Skipped or degraded image resources, skipped raw HTML comments when raw comment import is disabled or comments are empty, skipped embedded media/widgets, blocked stylesheet resources, disabled stylesheet links, missing stylesheet link `href` attributes, failed stylesheet HTTP statuses, stylesheet transport failures, stylesheet resource timeouts, rejected stylesheet content types, unsupported CSS declarations, unsupported mapped CSS values, and opt-in accessibility warnings populate `Diagnostics` and invoke `DiagnosticHandler`. |
| Structural limits | Supported | `MaxHtmlNodes`, `MaxHtmlDepth`, `MaxCssBytes`, `MaxTotalCssBytes`, and `MaxTableCells` stop conversion with `HtmlConversionLimitException` and error diagnostics. `MaxCssBytes` is applied before parsing inline/embedded/configured CSS and before fully buffering local or remote external CSS where source length is available. `MaxTotalCssBytes` bounds CSS processed across one conversion. |
| Stylesheet URI policy | Supported | `AllowedStylesheetUriSchemes` and `AllowedStylesheetHosts` reject disallowed external stylesheet sources before they are read or fetched and emit source-aware policy diagnostics. |
| Stylesheet content-type validation | Supported | `ValidateStylesheetContentTypes` and `AllowedStylesheetContentTypes` reject disallowed declared remote stylesheet media types before content is read. Missing content types are accepted for compatibility. |
| CSS aggregate byte budget | Supported | `MaxTotalCssBytes` rejects configured, embedded, local, and remote CSS that would exceed the conversion budget with `CssTotalSizeLimitExceeded`. |
| Stylesheet parse cache | Supported | Parsed stylesheet rules are cached by CSS content hash, reusing identical CSS across conversions without reusing stale rules when a local path or remote URL returns changed content. |
| Supported non-image aggregate resource budget | Supported | CSS is the supported external non-image resource family and is bounded per stylesheet and across the operation by `MaxCssBytes` and `MaxTotalCssBytes`. Executable and embedded media (`iframe`, `object`, `embed`, `video`, `audio`, and `canvas`) are rejected before resource loading, so they do not create an unbudgeted fetch path. |
| Parallel prefetch | Supported | Remote image candidates are deduplicated and prefetched with bounded concurrency while preserving deterministic projection order. |

## Table Conversion

| Feature | Import | Export | Notes |
| --- | --- | --- | --- |
| Basic tables | Supported | Supported | Rows and cells convert in both directions, including table-only HTML imports without leading empty placeholder paragraphs. |
| Header/body/footer row groups | Supported | Supported | Import preserves `thead` rows as repeating Word header rows and `tfoot` intent as last-row conditional formatting where Word can express it; export emits leading Word header rows as `<thead>` / `<th scope="col">`, body rows as `<tbody>`, and a final conditionally formatted row as `<tfoot>`. Multiple-row footer groups remain lossy because Word exposes only a last-row table-look flag. |
| Captions | Supported | Supported | Import supports table caption placement above or below the table; export emits adjacent `Caption`-style table captions as semantic `<caption>` elements. |
| Column groups | Supported | Partial | Import uses `colgroup`/`col` widths. Export emits Word column widths as `colgroup`/`col` markup when `IncludeTableColumnGroups` is enabled. |
| Colspan/rowspan | Supported | Supported | Grid spans and vertical merges convert in common and irregular table layouts. |
| Widths, borders, and spacing | Supported | Supported | Common CSS width, border, and table cell-spacing values are mapped; full browser border conflict resolution is outside the current contract. |
| Cell styles | Supported | Supported | Background color, horizontal alignment, vertical alignment, width, and border styles are handled for common tables. |
| Nested tables | Supported | Supported | Direct nested tables preserve their order relative to cell paragraphs and remain inside the preceding list item when that is their Word structure. Recursive export is bounded by `MaxTableNestingDepth`; malformed or deeper input is rejected instead of flattened silently. |

## List Conversion

| Feature | Import | Export | Notes |
| --- | --- | --- | --- |
| Ordered and unordered lists | Supported | Supported | Native Word numbering is used where possible. |
| Nested lists | Supported | Supported | Nested levels are tracked through Word list levels. |
| Block children in list items | Supported | Supported | Import preserves source order for paragraph and table blocks. Export keeps a direct nested table after a numbered paragraph inside that `li`, retains later list items in order, and uses Word numbering as the structural boundary. Arbitrary browser layout inside a list item is not implied. |
| Starts and continuation | Supported | Supported | `start`, `value`, reversed lists, and continued numbering are handled for common cases. |
| CSS list styles | Supported | Supported | Common bullet and ordered styles map to native OpenXML formats. |
| International ordered styles | Supported | Supported | Russian, Hebrew, Arabic, Hiragana, Katakana, and related OpenXML-backed formats are mapped when available. |
| Custom marker glyphs | Supported | Supported | Common editor markers such as dash, en dash, em dash, asterisk, and plus bullets are mapped, and arbitrary quoted CSS marker strings round-trip as custom Word bullet marker text. |
| Task list checkboxes | Supported | Supported | Checkbox task-list HTML imports as Word checkbox controls and exports back as disabled HTML checkboxes. Other native Word form controls export through the Word-to-HTML form-control path. |

## Word To HTML Export

| Area | Status | Current behavior |
| --- | --- | --- |
| Document shell | Supported | Emits HTML, root `lang` when the Word document language is set, head metadata, title, optional default CSS, additional meta tags, additional link tags, and optional body font family. |
| Paragraphs and headings | Supported | Emits `p`, `h1` through `h6`, `blockquote`, `hr`, paragraph classes, alignment, spacing, indentation, background, and border styles. |
| Definition lists | Supported | Emits consecutive imported or marker-styled definition term/description paragraphs as `dl`, `dt`, and `dd` instead of degrading indented descriptions to `blockquote`; empty internal marker paragraphs created by table-cell imports are skipped. |
| Runs | Supported | Emits `strong`, `em`, `u`, `s`, `del`, `ins`, `sup`, `sub`, `q`, `cite`, `dfn`, `time`, `abbr`, `mark`, `code`, font spans, run classes, colors, highlights, and run-level `lang` attributes when the run language differs from the document language. Imported `time datetime` values are preferred over parsing visible labels when available, and imported `mark` runs prefer semantic tags over generic highlight CSS. |
| Ruby text | Supported | Native Word ruby runs emit ordered `ruby`/`rb`/`rt` markup. HTML imports with explicit one-to-one segment pairs produce separate native Word ruby runs so the pairs survive export and package reload. |
| Links and bookmarks | Supported | Exports external and anchor hyperlinks as `a href`; ordinary Word bookmarks export as HTML `id` anchors and structural bookmarks export as matching structural elements. |
| Images and SVG | Supported | Emits `img` with base64 data URIs or file/external references, preserving image `alt` and `title` metadata where available; SVG can be emitted inline or as image references depending on options. |
| Lists | Supported | Emits `ol`, `ul`, `li`, `start`, `type`, optional CSS list styles, and optional reusable `word-list-*` definition classes plus a head stylesheet. |
| Tables | Supported | Emits tables with spans, widths, borders, background, horizontal and vertical alignment, cell spacing as `border-spacing`, cell styles, optional `colgroup` column widths, adjacent `Caption`-style table captions as `<caption>`, leading header rows as `<thead>` / `<th scope="col">` where Word marks them as repeating header rows, and a final `<tfoot>` row when Word last-row conditional formatting is enabled. |
| Figures and captions | Supported | Exports common image-plus-caption patterns and imported single-block figures, including a paragraph or table with a leading or trailing caption, as `figure` and `figcaption`. A figure with multiple independent content blocks remains readable but emits `HtmlFigureStructureFlattened` because Word has no native grouping carrier for that HTML shape. |
| Notes | Supported | Emits footnote and endnote sections when enabled. Exported `section.footnotes` and `section.endnotes` links import back as native Word footnotes and endnotes. Imported blockquote citation notes are suppressed during HTML export when the original `blockquote cite` attribute is restored, including after saved DOCX package reload when citation note metadata is present. |
| Comments | Supported | When `ExportComments` is enabled, emits linked Word comment references and a `section.comments` list with author, initials, date, and nested reply metadata. Exported comment sections import back as native Word comments with nested replies. Raw HTML comments are skipped by default with `HtmlCommentSkipped`; `ImportHtmlComments` imports non-empty raw comments with configured author metadata. Exact arbitrary comment-range reconstruction is not claimed. |
| Headers and footers | Supported | When `ExportHeadersAndFooters` is enabled, emits non-empty section headers and footers as semantic `header`/`footer` regions with section indexes and default/first/even type metadata. Those regions import back into the matching native section parts. Unsupported rich drawing content remains subject to normal conversion diagnostics. |
| Form/content controls | Supported | Emits native Word checkbox, structured text, dropdown list, combo box, and date picker content controls as disabled HTML form controls with selected values, option lists, and available metadata. Single-line structured text exports as `input type="text"`; multiline structured text exports as `textarea` so HTML import can round-trip line breaks and metadata. |
| Sections/page setup | Supported | When `IncludeSectionMetadata` is enabled, export wraps each Word section in `section.word-section` with page size, orientation, margins, and Word-like CSS dimensions. Exported metadata recreates native section boundaries and page settings; CSS layout is evidence, not a claim of browser/Word pagination identity. |
| Properties, definitions, and accessibility metadata | Supported | Built-in and optional typed custom properties export as metadata. Document/run language, image description/title, table accessible name/description, paragraph/run style definitions, and optional list definitions round-trip through their mapped HTML carriers. HTML `aria-label`/`aria-labelledby` and `aria-description`/legacy `summary` map to native Word table caption/description properties. Metadata with no Word carrier remains diagnostic or outside the documented profile. |
| Conversion diagnostics | Supported | `ToHtmlResult()` reports stable diagnostics when tracked revisions or fields are flattened, comments/headers/footers/notes are disabled, section layout is simplified, or macro/signature package content cannot exist in HTML. Use `RequireNoLoss()` when those diagnosed projections must block a workflow. |
| Export budgets | Supported | `MaxDocumentElements`, `MaxEmbeddedImageBytes`, `MaxTotalEmbeddedImageBytes`, and `MaxOutputCharacters` bound document traversal, individual and aggregate embedded-image reads, and serialized output. A breached limit raises `HtmlConversionLimitException` and records a stable error diagnostic. |

## Public Options

| Option surface | Status | Key controls |
| --- | --- | --- |
| `HtmlToWordOptions` | Supported | Named profiles through `CreateOfficeIMOProfile()`, `CreateUntrustedHtmlProfile()`, and `CreateTrustedDocumentProfile()`; cloneable reusable templates; font family, quote text, default page size/orientation, class styles, list styles, numbering continuation, heading numbering, base path, notes, raw HTML comment import and author metadata, image processing mode, `HttpClient`, resource timeout and concurrency, exact or highlight-palette text backgrounds, image limits, HTML/CSS/table limits, image and stylesheet content-type validation, image and stylesheet URI policy, diagnostics, opt-in accessibility diagnostics, unsupported CSS handling, stylesheet paths/contents, pre rendering mode, caption placement, and section handling. |
| Existing-container insertion | Supported | Parsed HTML can be appended through typed APIs to a document body, section header/footer, or `WordTableCell`; table-cell insertion preserves nested paragraphs, lists, tables, images, and package relationships. |
| `WordToHtmlOptions` | Supported | Font family, font styles, list styles, list definitions, paragraph/run classes, run color/highlight styles, paragraph spacing/indentation styles, footnote and endnote export, header/footer export, comment export, custom property metadata export, section metadata export, table column-group export, image embedding mode, additional meta/link tags, default CSS, and document/image/output budgets. `ToHtmlResult()` owns the output diagnostics. |
| `ReaderHtmlOptions` | Supported | `OfficeIMO.Reader.Html` exposes named adapter profiles for OfficeIMO-default, portable, and bounded untrusted HTML ingestion; nested `HtmlToMarkdownOptions` pass-through for markdown writer profiles, input limits, transforms, custom element converters, and visual round-trip hints; and `Clone()` for reusable registration templates. |

## Validation ownership

HTML-focused tests under `OfficeIMO.Html.Tests`, `OfficeIMO.Word.Tests`, and `OfficeIMO.Reader.Tests` cover import, export, tables, lists, images, links, inline styles, page settings, code blocks, task lists, diagnostics, resource limits, async and cancellation behavior, saved DOCX validity, and option contracts. The hashed [Word evidence corpus](../OfficeIMO.TestAssets/Documents/Word/EvidenceCorpus/README.md) combines ruby, forms, nested list/table blocks, single-block figures, native headers/footers/sections, and table accessibility metadata in one reciprocal fixture. `OfficeIMO.Word.Benchmarks` validates and measures field refresh, merge, structured comparison, package load, and Word-to-HTML workflows over deterministic 100- and 1,000-item corpora.
