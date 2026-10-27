# OfficeIMO visual design

This contract defines how the built-in `OfficeTheme` values become finished Word, Excel, PDF and HTML output. A document created without any styling must look deliberately designed: one type family per role, a clear heading hierarchy, restrained colour, tables that read at a glance and charts without decoration. The [API design](officeimo.api-design.md#themes) owns the theme model and the [visual quality](officeimo.api-design.md#visual-quality) release criteria; this document owns the values and the mapping rules.

[`Build/Visual/themes.json`](../Build/Visual/themes.json) is the machine-readable source for every value below. Its `palette`, `fonts`, `typeScale`, `tableStyle`, `chartStyle` and `spacing` objects map one-to-one onto `OfficeThemePalette`, `OfficeThemeFonts`, `OfficeThemeTypeScale`, `OfficeThemeTableStyle`, `OfficeThemeChartStyle` and `OfficeThemeSpacing`; lengths are points and colours are `RRGGBB` literals or `OfficeColor` role names. `extensions` holds the remaining design values: line spacing, title, subtitle, heading and link colours, heading weight, the title rule and heading spacing map onto `OfficeThemeTextStyle`; `table` and `chart` extend `OfficeThemeTableStyle` (border mode, header rule, header weight, numeric alignment) and `OfficeThemeChartStyle` (axis line and text colours, gridline and line widths, bar gap, markers); `surface` and `rule` map onto `OfficeThemeSurfaces` together with the `derived` accent and status tints; `page` and `footer` map onto `OfficeThemePage` and `OfficeThemeFooter`. `checks` records the generated contrast and colour-vision results. The built-in themes are generated from this file by `OfficeIMO.Core/Themes/generate_theme_builtins.py`; change values here, regenerate the checks and the built-ins, and do not copy numbers into writers by hand.

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

OfficeIMO builds every slide from a small set of designed layouts. The layout is chosen from the slide's content, its geometry comes from a 12-column grid, and its colours, fonts and base sizes come from the theme. Nobody gives coordinates, sizes or colours to get a good slide; `-X`, `-Y`, `-Width`, `-Height`, `-Column` and `-Span` remain for deliberate placement.

[`Build/Visual/powerpoint-layouts.json`](../Build/Visual/powerpoint-layouts.json) is the machine-readable source for everything below. `Build/Visual/powerpoint-mockups.py` renders reference slides from that file and `Build/Visual/themes.json`, so the pictures cannot drift from the values.

### Slide and grid

The default slide is 16:9, 13.333 × 7.5 in (960 × 540 pt).

| Measure | Value |
| --- | --- |
| Columns | 12, with a 24 pt gutter |
| Side margins | theme `spacing.slideMarginLeft` / `slideMarginRight` (Modern 48 pt, Boardroom 60 pt, Minimal 40 pt) |
| Top margin | theme `spacing.slideMarginTop` (Modern 36 pt, Boardroom 44 pt) |
| Live area bottom | fixed at 488 pt, so the footer never moves |
| Rows | 12, with a 12 pt row gutter; row height is the remaining height divided by 12 |
| Footer band | 506–522 pt, inside the bottom margin |

A box at column `c` spanning `n` columns starts at `margin + (c − 1) × (column + gutter)` and is `n × column + (n − 1) × gutter` wide; rows work the same way. With the reference margins (60 pt sides, 44 pt top) a column is 48 pt and a row 26 pt, which are the absolute boxes listed in the JSON.

On 4:3 (10 × 7.5 in) the same grid references apply with 42 pt side margins and an 18 pt gutter. Display and section titles scale down slightly, tile rows hold at most three items, and tables allow six columns before shrinking.

### Type

Slide text uses the theme's heading and body fonts. Sizes are theme type-scale tokens multiplied for viewing distance, rounded to even points from 18 pt up:

| Style | Token × factor | Modern | Boardroom | Minimum |
| --- | --- | --- | --- | --- |
| Display (cover) | title × 1.9 | 54 pt | 60 pt | 40 pt |
| Section title | title × 1.5 | 42 pt | 48 pt | 32 pt |
| Slide title | title × 1.143 | 32 pt | 36 pt | 26 pt |
| Subtitle, lede, card heading | body × 1.818 | 20 pt | 20 pt | 16 pt |
| Body (lists, text) | body × 1.636 | 18 pt | 18 pt | 16 pt |
| Small body (cards, steps) | body × 1.364 | 15 pt | 15 pt | 12 pt |
| KPI value | title × 1.714 | 48 pt | 54 pt | 32 pt |
| Quote | title × 1.286 | 36 pt | 42 pt | 24 pt |
| Table body / header | body × 1.273 / 1.182 | 14 / 13 pt | 14 / 13 pt | 11 pt |
| Eyebrow (caps, tracked) | caption × 1.333 | 12 pt | 12 pt | 10 pt |
| Caption / footer | caption × 1.222 / 1.111 | 11 / 10 pt | 11 / 10 pt | 9 pt |

Titles, subtitles and display text balance their line lengths so no line ends with a single word. Hyphenation is off, and a line never breaks at a hyphen when fewer than four characters would move to the next line (`walk-up` stays together).

Numbers in KPI values, results panels, table cells and chart labels use lining, tabular figures. Boardroom's Georgia has only old-style figures, so in Boardroom those numbers are set in the body font at the same weight; headings and running text keep Georgia.

### Colour

Layouts name palette roles only: `background`, `text`, `textMuted`, `accent1`, `accent2`, and the status roles for change indicators. Derived roles are computed, never stored:

- `surface` is the theme's surface tint (cards, quote slides, map panels).
- `hairline` is the theme's rule colour (rules, tile borders, connectors).
- `ink` is the darker of text and background (Section background, image scrim); `onInk` is the other.
- `onAccent` is white when it reaches 4.5:1 on `accent1`, otherwise `ink`.

Tables use the theme table style unchanged: header fill and text, banding and the optional header rule.

### Header and footer

Every content layout shares one header:

- A 28 × 3 pt accent mark sits above the eyebrow line on every content slide. It is the deck's one constant signature.
- The eyebrow (row 1) is `-Eyebrow`, or the title of the nearest preceding Section slide.
- The title starts on row 2, cols 1–10, at most two lines.
- The subtitle follows the title's last line by 10 pt.
- Content starts on the first row at least 36 pt below the last header line. Every slide keeps the same gap under its header, and content always starts on a row line.

The footer shows the deck title (or `PptFooter`) on the left and the slide number on the right, in the footer style. Title, Section and Closing slides have no footer; Quote shows only the number; ImageFull draws both in white over the scrim. Footer text never wraps, and it is written to the native footer and slide-number placeholders. `PptSlideNumber -IncludeTotal` prints `9 / 24`; `-NoFooter` and `-NoSlideNumber` on `New-OfficePowerPoint` or `PptSlide` turn them off for the deck or one slide.

### Layouts

| Layout | Content | Composition |
| --- | --- | --- |
| `Title` | title, subtitle, eyebrow, `-Author` and `-Date`; optional brand image | `accent1` background. Title stack in cols 1–9 anchored to row 10; a thin vertical rule at col 10 with the date and author beside it. |
| `Section` | title, subtitle | `ink` background. Title stack in cols 1–8; a 9 % two-digit section numeral fills cols 8–12 and bleeds off the bottom. |
| `TitleAndContent` | one list, text, table, chart or image | Lists and text keep a 60–70 character measure (cols 1–8); tables, charts and images use all 12 columns. A paragraph before the block becomes its lede. |
| `TwoContent` | two blocks | 6 + 6 columns; a text block beside a visual gets 4 columns and the visual 7, with one empty column between. |
| `Comparison` | two `PptColumn` with titles | Column headings with a 3 pt rule in `accent1` and `accent2`, content below. |
| `Kpi` | 1–6 `PptKpi`, optional takeaway paragraph | Equal tiles on one row (5 tiles divide the width in fifths; 6 make two rows of 3). Each tile has a top rule, the value, label and a coloured change indicator; the takeaway sits under a full-width hairline. |
| `Cards` | 2–6 `PptCard` | Surface panels with a top accent rule and index number. Title and body sit on a shared line across the row, so tall cards read as deliberate rather than half-empty. |
| `Process` | 2–8 `PptStep`, optional takeaway | Numbered circles joined by a hairline connector; the first step is filled. Five steps fit one row; six to eight make two rows. |
| `Quote` | one `PptQuote` with `-Attribution` | `surface` background, a large accent quote mark, the quote in cols 2–10 and a short rule above the attribution. |
| `ImageFull` | one photograph | Full-bleed cover crop around `-FocalPoint` (default centre) with a bottom and left `ink` scrim; title stack in cols 1–8. |
| `Closing` | title, optional next-steps list (≤ 4) or contact paragraph | `accent1` background; title stack left, numbered next steps right, separated by fine rules. |
| `CaseStudy` | `PptCard` narrative sections and `PptKpi` results | Challenge, approach and outcome in cols 1–7; a surface results panel with up to three metrics in cols 9–12. |
| `Logos` | 3–12 image-only `PptCard` | Hairline tiles, four or six per row, logos contained and shown in greyscale. |
| `Coverage` | up to 18 `PptLocation`, optional map `PptImage` | A dotted map surface (or the given map) in cols 1–8 with numbered pins and the `-MapLabel` caption; the numbered location list in cols 9–12. A `PptImage` on a Coverage slide is the map. |

### Automatic layout selection

`-Layout` wins when given. Otherwise the first matching rule chooses the layout:

1. First slide of the deck with no content: `Title`.
2. Any `PptLocation`: `Coverage`.
3. `PptKpi` together with `PptCard`: `CaseStudy`.
4. Any `PptKpi`: `Kpi`.
5. `PptCard` items that all have `-Image` and no text: `Logos`; any other `PptCard`: `Cards`.
6. Any `PptStep`: `Process`.
7. Two `PptColumn`, both with `-Title`: `Comparison`.
8. A single image close to the slide's aspect ratio and at least 1600 px wide: `ImageFull`.
9. A `PptQuote`, or no title and one short paragraph: `Quote`.
10. Last slide of a deck of five or more, with no content or a short list: `Closing`.
11. Title only: `Section`.
12. Two blocks: `TwoContent`. One block: `TitleAndContent`.
13. Three or more blocks: the first two stay on a `TwoContent` slide and the rest continue on the next slide.

A block is a list, a run of paragraphs, a table, a chart, an image or a `PptColumn`. Notes and footers do not count. If content does not fit the requested layout, OfficeIMO reports `layout-content-mismatch` and selects automatically.

`Test-OfficePowerPoint` and `Get-OfficePowerPointSlide -Summary` report the deck's rhythm: more than three consecutive slides of one layout, more than three dense slides in a row, more than seven slides without a Section, or a deck of eight or more slides without a Closing slide. These are reported, never applied silently.

### Overflow

Content that does not fit is handled; nothing is clipped and nothing silently shrinks below its minimum.

**Titles.** A title that needs more than two lines is split at `: ` or a dash into title and subtitle when no subtitle was given. Otherwise it shrinks in 2 pt steps to 26 pt, then takes a third line and the content moves down one row (`title-long`). Titles are never truncated.

**Lists.** A list shrinks in 2 pt steps to 16 pt. If it still does not fit, it is split into the fewest slides that fit at the original size, with balanced item counts (14 items become 7 + 7, never 12 + 2), so every slide in the run uses one size. Splits fall between top-level items, and numbering continues.

**Tables.** Rows per slide are computed from the content zone and the table row height (typically 8–9 data rows at 16:9). A longer table splits into balanced slides with the header repeated, and `Rows 9–16 of 30` appears under each part. `-MaxRows` overrides the limit; `-NoSplit` keeps one slide and shrinks to 11 pt before reporting `table-overflow`. Columns are never split: more than eight columns shrink to 11 pt and then report `table-too-wide`, which names `-Property` as the fix.

**Tiles.** KPIs, cards and steps shrink their text to its minimum, then split into balanced slides above six KPIs or cards, eight steps, twelve logos or eighteen locations.

**Charts** never split. A column or bar chart with more than twelve categories becomes a horizontal bar chart when `-Type` was not given.

**Continuation slides** keep the eyebrow and title and drop the subtitle and lede. The title gains `(continued)`, or `(continued 2/3)` for runs of three or more, in muted regular text at subtitle size. If that suffix would add a title line, it moves to the eyebrow (`CHANGE ADVISORY BOARD · CONTINUED 2/3`) so continuation slides keep the first slide's header and capacity. `New-OfficePowerPoint -ContinuationFormat` changes the wording, for example `'{0} (cont. {1}/{2})'`.

### Decks from data

`Export-OfficePowerPoint` builds a complete deck from pipeline objects:

1. **Title** slide: `-Title` (default: the file name), `-Subtitle` (default: the item count) and the date.
2. **At a glance** (`Kpi`), when there are two or more rows and `-NoSummary` is not given: the item count; up to three groups of the group property (`-GroupBy`, otherwise `-Category`, otherwise the first text property with two to six distinct values) with their share coloured by status; and the total of the first numeric `-Value`. A takeaway names the members of groups with three or fewer items, such as `Outage: Printing. Degraded: VPN, Monitoring, Wi-Fi.`
3. **Table** slides (`TitleAndContent`), split as described above, unless `-NoTable`.
4. **Chart** slide, when `-Category` and `-Value` are given: `-ChartTitle` or `{Value} by {Category}`. Above twelve categories it shows the top twelve and one `Other (n)` bar in the hairline colour.

Every table slide carries speaker notes with its row range, the source type and the export time. Data exports have no Closing slide. The footer shows `-Title` and the slide number.

### Speaker notes

`PptNotes` sets the notes of its slide and never affects the layout. Continuation slides start without notes. Converting a Markdown outline turns blockquotes into notes.

### Corporate templates

When `-Theme` is a `.potx` or `.pptx`, the template's layouts replace the built-in geometry:

1. A template layout named like the OfficeIMO layout (`Kpi` or `OfficeIMO Kpi`) is used as it is.
2. Otherwise the layout type decides: Title → title, Section → section header, TitleAndContent → title and content, TwoContent → two content, Comparison → comparison, and the remaining layouts → title only.
3. Placeholders bind by type and then by index, and keep the template's position, fonts and colours.
4. Compositions without a template placeholder (KPI tiles, cards, process, coverage, logos) are drawn inside the mapped layout's body area with the template theme's fonts and colour scheme (dark 1 → text, light 1 → background, accents 1–6 → accents).
5. The 12-column grid is re-derived from that body area, so `-Column` and `-Span` keep working, and overflow limits are recomputed from it. Minimum sizes stay as defined here.
6. A missing layout is reported once per deck as `template-layout-missing`, never as a failure.

## Evidence

The mockups were generated from `themes.json` by script and rendered with headless Microsoft Edge using the Office font files. Each theme has a Word report page, an Excel sheet, a chart sheet with six series and a palette proof showing the simulated colour-vision deficiencies and greyscale. The mockups are design references; [acceptance](officeimo.api-design.md#acceptance) still requires renders of real output in Microsoft Office, LibreOffice and OfficeIMO's renderer.
