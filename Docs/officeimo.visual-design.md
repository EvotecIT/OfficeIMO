# OfficeIMO visual design

This contract defines how the built-in `OfficeTheme` values become finished Word, Excel, PDF and HTML output. A document created without any styling must look deliberately designed: one type family per role, a clear heading hierarchy, restrained colour, tables that read at a glance and charts without decoration. The [API design](officeimo.api-design.md#themes) owns the theme model and the [visual quality](officeimo.api-design.md#visual-quality) release criteria; this document owns the values and the mapping rules.

[`Build/Visual/themes.json`](../Build/Visual/themes.json) is the machine-readable source for every value below. Its `palette`, `fonts`, `typeScale`, `tableStyle`, `chartStyle` and `spacing` objects map one-to-one onto `OfficeThemePalette`, `OfficeThemeFonts`, `OfficeThemeTypeScale`, `OfficeThemeTableStyle`, `OfficeThemeChartStyle` and `OfficeThemeSpacing`; lengths are points and colours are `RRGGBB` literals or `OfficeColor` role names. `extensions` holds values the designs need that the current model has no field for, `derived` holds values writers compute from the palette, and `checks` records the generated contrast and colour-vision results. Change values there and regenerate the checks; do not copy numbers into writers by hand.

## Themes

| Theme | Fonts (heading / body) | Accent | Use it for |
| --- | --- | --- | --- |
| `Modern` (default) | Aptos / Aptos | Cobalt `3F64D7` | Status reports, exports, dashboards, anything generated without a brief |
| `Classic` | Cambria / Cambria; charts Calibri | Navy `1F3A5F`, ochre rule | Policies, proposals, contracts and letters that will be printed |
| `Minimal` | Arial / Arial | Graphite `353535`, vermilion second series | Technical notes, data appendices, output that must match in LibreOffice |
| `Boardroom` | Georgia / Aptos | Deep navy `14335C`, muted gold | Board packs, quarterly reviews, investor material |
| `Dark` | Aptos / Aptos | Sky `68A1F9` on `0F1620` | HTML reports, on-screen PDF, wallboards, presentations |

- **Modern** is airy and neutral. Headings use the text colour, not the accent; colour appears in the subtitle, bullets, the table header rule and charts. Tables have a 10 % accent-tinted header and light banding.
- **Classic** is a traditional serif document. Title and headings are navy, the title has an ochre rule beneath it, and tables have a solid navy header on warm banding. Chart text stays sans-serif.
- **Minimal** uses no fills. Tables have a black rule under a white header and hairlines between rows, without banding. Arial and Liberation Sans share metrics, so the layout is the same in Microsoft Office and LibreOffice.
- **Boardroom** uses regular-weight Georgia headings at larger sizes with a gold rule, a solid navy table header and generous margins. Data labels are on by default because executive readers quote numbers.
- **Dark** is for screens: slides, HTML and on-screen PDF. Word and Excel documents are printed, edited and extended by other people, so in those formats `Dark` keeps a light page and uses its accent palette and type; nothing is painted dark.

## Typography

Documents name fonts that must exist on the viewer's machine; PDF output embeds the font that was resolved. Every family below has an ordered fallback list in `themes.json`, used by PDF font resolution, the HTML CSS stack and the Word font table's alternate name.

| Primary | Where it exists | Fallback order | Linux/LibreOffice result |
| --- | --- | --- | --- |
| Aptos | Microsoft 365 on Windows (cloud font, downloaded on first use) and macOS (bundled with Office) | Calibri, Carlito, Helvetica Neue, Arial, Liberation Sans | Carlito, metric-compatible with Calibri |
| Cambria | Windows, Office for Mac | Caladea, Georgia, Liberation Serif, Times New Roman | Caladea, metric-compatible |
| Georgia | Windows, macOS | Gelasio, Cambria, Caladea, Liberation Serif | Gelasio (metric-compatible) when installed, otherwise Caladea |
| Arial | Windows, macOS | Liberation Sans, Arimo, Helvetica Neue, Helvetica | Liberation Sans, metric-compatible |
| Calibri (Classic charts) | Windows, Office for Mac | Carlito, Aptos, Helvetica Neue, Arial, Liberation Sans | Carlito |

The trade-off: `Modern`, `Boardroom` and `Dark` use Aptos because it is the current Office default and gives the best result for most readers. LibreOffice has no Aptos substitute, so those documents fall back to Carlito there; Carlito is narrower than Aptos, which means fallback text reflows shorter and never overflows a layout sized for Aptos. `Minimal` and `Classic` use fonts with metric-compatible open substitutes and should be chosen when Linux fidelity matters more than the current Office look. Aptos Display is not used: it is a separate cloud family that Office for Mac does not bundle. Aptos, Calibri, Cambria, Arial and Georgia allow editable embedding (OS/2 `fsType` 0x0008), so PDF output may embed subsets of them; OfficeIMO does not redistribute the font files.

Heading levels 4 to 6 differ from body text by weight and size only. Line spacing is 1.2 for `Modern` and 1.25 for the other themes; heading space before and after is listed per level under `extensions.headingSpacing`.

## Colour and contrast

Every text and background pair a theme specifies meets WCAG 2.2 AA: 4.5:1 for text, 3:1 for chart marks against their background. The checked pairs and their ratios are in each theme's `checks.contrast`: body and muted text on the background and on banded rows, heading, subtitle and link colours, table header text on header fill, axis text, every status colour on the background and on its own tint, body text on each status tint, and every series colour. The lowest text ratio is muted text in `Modern` at 5.89:1; the lowest chart mark is `Minimal` series 4 at 3.10:1. Gridlines and table borders are deliberately below 3:1. They are structure, not information, and are recorded under `checks.decorative`.

### Series palette

The series palette is `Accent1` to `Accent6` as theme roles, so charts follow a theme edited later in Office. Each palette was built on a CIE L* ladder, so every series has a distinct lightness, and then checked with Machado (2009) full-severity protanopia, deuteranopia and tritanopia simulation:

| Theme | Normal | Protanopia | Deuteranopia | Tritanopia | Greyscale min ΔL* |
| --- | --- | --- | --- | --- | --- |
| Modern | 22.7 | 13.2 | 13.8 | 12.6 | 6.8 |
| Classic | 21.2 | 12.7 | 14.7 | 12.1 | 6.8 |
| Minimal | 21.6 | 14.0 | 11.8 | 13.6 | 6.8 |
| Boardroom | 22.3 | 12.0 | 11.0 | 14.9 | 6.2 |
| Dark | 25.9 | 13.7 | 14.7 | 15.1 | 7.0 |

Values are the minimum CIEDE2000 difference over all series pairs; above about 10 the colours are clearly distinct. In greyscale the 3:1 requirement limits light themes to L* 22–61, so six series can be at most about 7 L* apart. Neighbouring series alternate between dark and light to maximise adjacent contrast. Line charts with more than four series, and any chart intended for monochrome print, also use dash patterns in series order: solid, dash, dot, dash-dot, long dash, long dash-dot. A seventh series reuses series 1 with the second dash pattern.

### Semantic roles

`Success`, `Warning` and `Danger` are shared by the light themes (`08684A`, `865500`, `BF3B2C`) and brightened for `Dark` (`3FA66E`, `F2C14E`, `FF8F8F`). Each passes 4.5:1 on the background and on its tint.

- **Text:** a status word ("Done", "At risk", "Below plan") uses the status colour in bold. Never colour a whole sentence.
- **Table rows:** a highlighted row uses the status tint (`derived.*Fill`, the status colour at 12 % over the background, 18 % in `Dark`) behind normal text, and only the status cell uses the status colour. Use at most one highlight per five rows; when everything is highlighted, nothing is.
- **Excel:** status columns use the status font colour and no fill; conditional formats that need fills use the same tints.
- **Never colour alone:** text-grade status colours stay close for red-green deficiencies (Warning and Danger are the closest pair, ΔE 6.5 under protanopia) and nearly identical in greyscale, so a status always carries a word or symbol (▲ ▼ ●). Status colours are not series colours; a chart that encodes status uses the status roles with labels.

## Word

- **Styles:** `Title` (title size, heading font, `titleColor`), `Subtitle` (Heading 3 size, `subtitleColor`), `Heading 1`–`Heading 6` (type scale, heading font, `headingColor`, bold unless `headingBold` is false, `keepNext`), `Normal` (body size, body font, Text, `paragraphAfter`, theme line spacing), `Caption` (caption size, TextMuted, no italics), `List Paragraph` (`listIndent`, half `paragraphAfter` between items, bullets in Accent1). Themes with `titleRule` add a bottom border to `Title`, or to `Subtitle` when one follows. The theme part receives the palette in its native slots and the fonts as major/minor fonts, so Word's own theme pickers stay consistent.
- **Tables:** one generated table style per theme, named `OfficeIMO <Theme>`, with the header row (`HeaderFill`, `HeaderText`, bold, repeated on each page), banding from `BandFill`, horizontal borders only (`BorderColor`, `BorderWidth`), no vertical borders, `CellPadding` top and bottom and 1.4× that left and right, and the header rule from `extensions.table.headerRule`. Numeric columns are right-aligned. Body text in tables is one point smaller than `Normal`.
- **Header and footer:** footer size and colour come from `extensions.footer`, with an optional hairline rule above. The footer is one paragraph with a right-aligned tab stop at the text width: the document title on the left, then `Internal · Page {PAGE} of {NUMPAGES}` after the tab. Fields are inline runs in that paragraph, never separate paragraphs, so the line cannot split.
- **Charts:** native charts with `chartStyle` values (see [Charts](#charts)), sized to the text width and followed by a `Caption`.
- **Page:** margins from `extensions.page`. `Dark` keeps a white page (see [Themes](#themes)).

## Excel

- **Cell styles:** `Normal` uses the body font and size; `Title` and `Heading 1`–`4` follow the type scale; `Total` uses bold with a 1 pt Text top border. The workbook theme receives the palette and fonts.
- **Tables:** `AddTable` and `Export-OfficeExcel` create a custom table style from `tableStyle`: header fill and text, banded rows, horizontal hairlines and no vertical lines. `Minimal` turns banding off and uses a 1 pt Text rule under the header.
- **Header row:** bold, frozen (pane below the header), AutoFilter on, height taken from `CellPadding` plus the body size.
- **Number formats from data:** integers `#,##0`; non-integer numbers `#,##0.00`; dates `yyyy-mm-dd`; date-times `yyyy-mm-dd hh:mm`; timespans `[h]:mm:ss`; booleans as `TRUE`/`FALSE`. Percent (`0.0%`) and currency (`$#,##0` with the culture's symbol) are applied only when declared or when the column header ends in `%` or names a currency code; values are never rescaled. Negative numbers use a leading minus in Text; `Danger` is reserved for status columns.
- **Column fitting:** fit to the widest of the header and the first 1,000 values plus two characters of padding, clamp to 8–60 characters and wrap headers longer than that.
- **Charts:** placed beside the table, aligned to the table's header row, with `chartStyle` formatting and no chart-area border in light themes.
- **Dark:** keeps Excel's light grid and applies only `Dark`'s accents and series palette. Filling cells to imitate a dark page breaks printing, fights the grid and surprises anyone who adds rows.

## PDF

The composition builder takes its defaults from the theme. Fonts are resolved through the fallback lists and embedded as subsets; type sizes, heading spacing and `paragraphAfter` match Word, so a document converted to PDF and the same document composed directly look the same. Tables, headers and footers follow the Word rules, including the single-line `Page N of M`. Page margins come from `extensions.page`; `Dark` paints the page background.

## HTML

Each theme becomes CSS custom properties: `--office-bg`, `--office-text`, `--office-muted`, `--office-accent-1`…`6`, `--office-success|warning|danger` and their `-fill` tints, the font stacks (primary plus fallbacks, ending in a generic family), the type scale in `pt`, and table and chart tokens. `Dark` sets `color-scheme: dark`. Tables use `font-variant-numeric: tabular-nums` and right-align numeric cells.

## Charts

These rules apply in every format:

- Text is sans-serif: the `axisFont` and `titleFont` fonts, even in serif themes. The title is the chart title size in semibold Text; the axis labels are the axis size in TextMuted.
- Gridlines are horizontal only, `gridlineColor` at 0.5 pt. There is no plot-area border or fill, the value axis has no line, and the category axis has a single line in `extensions.chart.axisLine`.
- Charts are never 3-D, and have no shadows, gradients, bevels or rounded columns by default. Doughnut and pie charts are allowed only for two to five parts of a whole.
- Columns use `barGapPercent` between groups, with no gap inside a group. Lines use `lineWidth` with round joins; markers appear only where `markers` says so, and always on series with fewer than three points.
- The legend sits top-left under the title, in series order, and is omitted for a single series. A line chart with three or fewer series may label line ends directly instead.
- With `showDataLabels`, a column chart labels every column when there are 24 or fewer, and a line chart labels only the last point of each series. Labels use the axis size in Text.
- Value axes start at zero for columns and bars, and use round steps (1, 2, 2.5 or 5 × 10ⁿ) with four or five intervals. Tick labels use the data's number format with the fewest decimals that keeps them distinct. Units go in the title ("USD thousands"), not on every tick; compact suffixes (K, M) are for dashboards only.

## PowerPoint layouts

Slide layouts and their placeholders are defined in a separate section ([`Build/Visual/powerpoint-layouts.json`](../Build/Visual/powerpoint-layouts.json)).

## Evidence

The mockups were generated from `themes.json` by script and rendered with headless Microsoft Edge using the Office font files. Each theme has a Word report page, an Excel sheet, a chart sheet with six series and a palette proof showing the simulated colour-vision deficiencies and greyscale. The mockups are design references; [acceptance](officeimo.api-design.md#acceptance) still requires renders of real output in Microsoft Office, LibreOffice and OfficeIMO's renderer.
