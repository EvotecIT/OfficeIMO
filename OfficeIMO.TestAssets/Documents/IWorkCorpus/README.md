# iWork reader corpus

These package fixtures prove the bounded Pages, Numbers, and Keynote reader against files produced by multiple iWork generations. They remain read-only test inputs. Separate OfficeIMO-created native Keynote, Pages and Numbers fixtures qualify their bounded creation models.

## Large native Keynote reference

The [512-slide reference](native-exports/keynote-scale-512-v15.4.json) is an Apple
Keynote 15.4 derivative of the licensed two-slide color fixture. Native duplication
produces 256 repetitions of two text layouts; save, close/reopen and inspection
reach slide 512 without repair. Independent decoding checks every ordered native
slide identity, text box, note, background and frame against Apple's PPTX export.
All 512 PDF pages pass complete text and canvas checks. Six representative Cairo
renders are inspected; Poppler Splash rendering and bbox ligature extraction have
the recorded limits. This is a repeated-layout scale reference, with broader
object diversity and complete appearance outside its qualification.

Reproduce the oracle in isolated evidence environments with pypdf 6.10.0,
numbers-parser 4.19.0 and the manifest's pinned Keynote descriptor:

```sh
python Build/IWork/extract-keynote-scale-pdf.py <apple.pdf> <pdf-text.json>
pdftotext -bbox -enc UTF-8 <apple.pdf> <pdf-bbox.xhtml>
python Build/IWork/extract-keynote-scale.py <native.key> --descriptor <keynote-descriptor.pb> --pptx <apple.pptx> --pdf <apple.pdf> --pdf-text <pdf-text.json> --pdf-bbox <pdf-bbox.xhtml> --out <decode.json>
```

The optional `--destination-pptx` argument checks all reconstructed OfficeIMO
slides through independent XML reading. The benchmark's `NativeLarge` profile
uses this source and verifies every unit outside measurement. No oracle tools
enter OfficeIMO's runtime or normal build requirements.

| Folder | Upstream | Revision | Producer evidence | License |
|---|---|---|---|---|
| `nim-iwork` | [nim-iwork](https://github.com/halcyon-oss/nim-iwork) | `60a8e875692ac934956cdb5b39f88e557c97cca4` | Pages, Numbers, and Keynote 14.5 fixtures | MIT, copyright Alfred |
| `iwork-converter` | [iwork-converter](https://github.com/obriensp/iwork-converter) | `be4828260466c9c022ec12f9dc9dbfeb15ab1dea` | Pages 14.1, Numbers 11.1, and Keynote 8.1 fixtures | MIT, copyright Steve Dunham |
| `numbers-parser` | [numbers-parser](https://github.com/masaccio/numbers-parser) | `1c6c5c3d2e29a9abb601596678089f0a6c85d64c` | Numbers 15.1 fixture plus independently produced formula and merged-range workbooks | MIT, copyright Jon Connell |
| `picodocs` | [PicoDocs](https://github.com/PicoMLX/PicoDocs) | `5c18743d3d8120a76da124bd512a3cf5bcc28e82` | Pages 14.4.1 package produced by Pages 14.5 with sections, headers, footers, an image, hyperlinks, and three editable tables | MIT, copyright Pico MLX |
| `keynotekit` | [KeynoteKit](https://github.com/memfrag/KeynoteKit) | `5e01f1e061b608e16c4480444d8e04790a625b34` | Keynote 15.2.1 image and editable-table fixtures independently maintained as parser/writer regressions | 0BSD, copyright Martin Johannesson |

The complete upstream license notices are reproduced in `OfficeIMO.IWork/THIRD-PARTY-NOTICES.md`. Fixture provenance and expected semantic assertions live beside the executable corpus tests in `OfficeIMO.IWork.Tests`.

The parameterized independent-corpus output check converts all 12 fixtures listed by `IWorkCorpusTests.Corpus` to editable DOCX, XLSX or PPTX, saves and reopens the result, and validates the destination Open XML schema. It explicitly accepts partial reconstruction and normalizes worksheet names. This gate catches invalid destination packages; separate semantic tests establish supported content. Passing it does not qualify native application rendering, pagination or complete document fidelity.

## Native Numbers scalar-table creation

The [Numbers creation manifest](native-exports/numbers-created-v14.5.json) pins a template-free OfficeIMO two-sheet/three-table spreadsheet, its native Numbers 14.5 numeric edit/save/close/reopen, and PDF/XLSX exports. `IWorkNumbersWritingTests.Example()` reproduces the frozen input with scalar text, finite numbers, booleans, Polish text, emoji, an inline LF, literal `=` text and a completely empty 2-by-3 table. Both native PDF pages are inspected. These are newly authored MIT fixtures; no seed archive or standalone font is distributed.

`numbers-created-graph-v14.5.json` independently decodes all 108 objects with complete required fields and object references, typed scalar tiles/catalogs, style networks, ordered sidebar/sheet tables, modern custom-format references and mandatory blank header/footer storages. Reproduce it with the manifest's pinned protobuf descriptor and isolated numbers-parser version:

```sh
python Build/IWork/verify-numbers-creation.py OfficeIMO.TestAssets/Documents/IWorkCorpus/native-exports/numbers-created-v14.5.numbers --descriptor <numbers-descriptor.pb> --out <graph.json>
```

Native readback and strict saved/reopened OfficeIMO XLSX retain all tables and typed values after B2 changes from 12.5 to 20.75. The populated tables agree with Apple's XLSX export. Apple's empty-grid positioning is B2:D3; the adapter produces a table-relative A1:C2 grid. Exact Excel positioning and appearance, broader authored objects, formulas, formats, comments, other native versions and fonts/scripts remain unqualified.

## Native source preservation

The [native preservation manifest](native-exports/native-preservation-macos.json) records unchanged saves of the Apple-edited Pages body, Numbers scalar tables and Keynote image deck, plus a preservation-aware translation of Keynote image record 154. These newly authored MIT fixtures retain the original resources and unknown content; no external font is redistributed.

The independent verifier compares every file-entry payload for unchanged saves and independently decompresses all IWA archives. For the translation it checks the complete raw archive against the original after replacing only the two fixed32 position values, including exact ArchiveInfo/MessageInfo bytes and auxiliary/sibling payloads. Pinned descriptors additionally check known and unknown protobuf fields. Reproduce the translation comparison with:

```sh
python Build/IWork/verify-native-preservation.py OfficeIMO.TestAssets/Documents/IWorkCorpus/native-exports/keynote-created-images-edited-v15.4.key OfficeIMO.TestAssets/Documents/IWorkCorpus/native-exports/keynote-preserved-translated-v15.4.key --descriptor <numbers-descriptor.pb> --image 154 100 130 --out <evidence.json>
```

Omit `--image` and `--descriptor` to compare a reproduced unchanged encoding. The manifest pins source/output hashes and record/entry counts for all three applications. Byte comparison applies to OfficeIMO output before native opening; Keynote can update application view state during inspection.

The [direct image resize manifest](native-exports/keynote-preserved-resized-v15.4.json) extends that proof to composed moves and resizes of opaque PNG, alpha PNG and JPEG images. Keynote 15.4 opens all three slides without repair, changes the JPEG width from 420 to 421 points and retains that edit after save/close/reopen. Independent native-graph and PPTX decoding checks exact image bytes, descriptions and final frames. Native and composed PDF inspection checks all three slides, painted image bounds, transparency samples and stacking; typography retains the recorded font-metric differences. Reproduce the raw IWA comparison with:

```sh
python Build/IWork/verify-native-preservation.py OfficeIMO.TestAssets/Documents/IWorkCorpus/native-exports/keynote-created-images-edited-v15.4.key OfficeIMO.TestAssets/Documents/IWorkCorpus/native-exports/keynote-preserved-resized-v15.4.key --descriptor <keynote-descriptor.pb> --image-frame 125 80 125 420 168 --image-frame 139 80 135 360 180 --image-frame 154 100 135 420 168 --out <evidence.json>
```

The descriptor hash and isolated numbers-parser/protobuf versions are pinned in the manifest. Only the requested fixed32 position/size spans differ; all raw headers, sibling records and other file-entry payloads remain intact. Native visual caches remain unassessed and require explicit acceptance. Other producer versions, crop/mask/rotation states, nested-index mutations and broader semantic editing retain their separate gates.

Keynote 15.4 reports the requested X100/Y130 frame, accepts a native one-point horizontal edit and retains X101/Y130/W600/H240 after save/close/reopen. All three native PDF pages are inspected, and saved-source resources match the original PNG/JPEG payloads. Unchanged outputs open in Pages 15.4, Numbers 14.5 and Keynote 15.4, retaining the observed two-page body, two-sheet/three-table workbook and three-slide geometry. This is fixture-specific native acceptance; unrefreshed preview/cache behavior, wider native metadata and other producer versions remain unassessed.

The [Numbers scalar-edit manifest](native-exports/numbers-preserved-scalars-v14.5.json) qualifies six existing numeric/boolean replacements across two tables in a two-sheet workbook. Numbers 14.5 opens the output without repair and retains a further Data B2 edit from 42.125 to 43.125 after save/close/reopen. Independent Numbers/XLSX decoding agrees on all typed values. The native PDF covers both sheets, and both pages are rendered and inspected. The raw verifier checks 103 entries and 316 records; only the six requested fixed-width scalar spans in tile records 150/175 differ:

```sh
python Build/IWork/verify-native-scalar-preservation.py OfficeIMO.TestAssets/Documents/IWorkCorpus/native-exports/numbers-created-edited-v14.5.numbers OfficeIMO.TestAssets/Documents/IWorkCorpus/native-exports/numbers-preserved-scalars-v14.5.numbers --descriptor <numbers-descriptor.pb> --cell 141 2 2 number 42.125 --cell 141 2 3 boolean false --cell 141 3 2 number -7.5 --cell 141 3 3 boolean true --cell 166 1 2 number 123.75 --cell 166 2 1 boolean true --out <evidence.json>
```

The descriptor and decoder versions are hash-pinned. Native visual caches remain unrefreshed and require explicit acceptance. This profile admits Decimal128 numbers and native booleans in active BNC storage with empty legacy placeholders; formulas, insertion, type changes, shared tiles and nested-index mutations are rejected. XLSX padding loss and the resulting complete-PDF rejection remain explicit. Other versions, fonts and broader edits require separate qualification.

The [Pages paragraph-edit manifest](native-exports/pages-preserved-paragraphs-v15.4.json) qualifies three fixed-length replacements in an Apple-saved layout document. Pages 15.4 opens the output without repair, changes the final paragraph again and retains that change after save/close/reopen. Native DOCX/PDF, strict OfficeIMO DOCX and complete composed PDF cover the authored body, spacing, tabs and untouched rich paragraph. Reproduce the raw comparison of all 13 file entries and 93 IWA messages with:

```sh
python Build/IWork/verify-native-text-preservation.py OfficeIMO.TestAssets/Documents/IWorkCorpus/native-exports/pages-created-layout-edited-v15.4.pages OfficeIMO.TestAssets/Documents/IWorkCorpus/native-exports/pages-preserved-paragraphs-v15.4.pages --descriptor <pages-descriptor.pb> --changes OfficeIMO.TestAssets/Documents/IWorkCorpus/native-exports/pages-preserved-paragraphs-v15.4-changes.json --out <evidence.json>
```

Only the requested UTF-8 ranges differ; native text properties, offsets, headers, siblings and all other entry payloads remain identical. Both UTF-16 and UTF-8 lengths are retained. Native caches remain unrefreshed, font substitution affects composed appearance, and broader text/structural edits require separate evidence.

## Native Pages body creation

The [inline scalar-table manifest](native-exports/pages-created-tables-v15.4.json) pins mixed body/table and table-only creation, independent graph decoding and Pages 15.4 acceptance. `IWorkPagesTableWritingTests.Example()` reproduces the mixed document. A numeric edit survives native save/reopen and DOCX export; the original PDF and screenshots qualify sequential placement. These are newly authored MIT fixtures. Native typography, localized display, broader table formatting and receiving-application DOCX row sizing remain separate limits.

The [Pages creation manifest](native-exports/pages-created-v15.4.json) pins a template-free OfficeIMO body document, its native Pages 15.4 edit/save/close/reopen, and PDF/DOCX exports. `IWorkPagesWritingTests.Example()` reproduces the frozen 83-paragraph input with Arial and Times New Roman, independent character decorations, explicit plain resets, a blank paragraph, Polish text and a non-BMP emoji. The original graph, native-saved edit, all two PDF pages and reference DOCX are newly authored MIT fixtures; no seed archive or standalone font is distributed.

`pages-created-graph-v15.4.json` records pinned-schema independent decoding of all 189 objects and their declared references. It includes the six required empty header/footer storages, even though headers/footers are hidden. Missing these objects can permit initial open but produce an unreadable package after native save. Reproduce the graph with:

```text
python Build/IWork/verify-pages-creation.py OfficeIMO.TestAssets/Documents/IWorkCorpus/native-exports/pages-created-v15.4.pages --descriptor <pages-descriptor.pb> --out <graph.json>
```

The manifest identifies the schema revision, descriptor hash and isolated numbers-parser version. Native Pages/PDF retain the rich styling after the text edit. Apple's DOCX export agrees with the complete authored content and page dimensions/margins but changes selected font runs and bold/italic overrides; native DOCX typography is not qualified as equivalent. Strict OfficeIMO DOCX separately preserves source run properties and includes the empty paragraph projected from the final LF. Other native versions, fonts, scripts, pagination contexts and broader authored objects remain open.

## Native Pages default headers and footers

The [default header/footer manifest](native-exports/pages-created-headers-v15.4.json) pins a two-page body with one left header and footer, explicit 30/36-point distances, bold/color header styling and italic Unicode footer text. `IWorkPagesHeaderFooterWritingTests.Example()` reproduces the deterministic creation. Pages 15.4 edits both areas, saves and reopens without repair. Native PDF and strict OfficeIMO DOCX retain the styles; Apple's Word export retains text and distances while changing font and bold/italic declarations. Complete composed PDFs retain all body lines and repeat both areas on every page. Preview captures qualify this rendering; the manifest records pagination and headless rasterizer differences.

```sh
python Build/IWork/verify-pages-creation.py OfficeIMO.TestAssets/Documents/IWorkCorpus/native-exports/pages-created-headers-v15.4.pages --descriptor <pages-descriptor.pb> --page-text --out <graph.json>
python Build/IWork/verify-pages-header-footer.py OfficeIMO.TestAssets/Documents/IWorkCorpus/native-exports/pages-created-headers-edited-v15.4.pages OfficeIMO.TestAssets/Documents/IWorkCorpus/native-exports/pages-created-headers-edited-v15.4.docx --descriptor <pages-descriptor.pb> --out <native-edit.json>
```

Both decoders use the pinned schemas and isolated numbers-parser version recorded in the manifest. Center/right positions, mixed runs, alternate-page templates, fields, more sections and loaded-source page-text editing remain outside this authored profile.

## Native default paragraph layout

`native-exports/numbers-simple-v14.5.json` pins the unchanged `nim-iwork/simple.numbers` source and its Numbers 14.5 XLSX/PDF exports. Native body-cell controls show a single line, no custom tab stops, natural horizontal alignment, Helvetica Neue 10-point text, and wrapping. Independent numbers-parser 4.19.0 schemas identify empty `12/13` line-spacing and `12/25` tab declarations in the selected body/header styles. The XLSX oracle adds one table-title row; comparisons offset that row and check all nine types, values, font sizes, header emphasis and General alignment. Empty shared declarations reset inherited relative spacing/custom tabs; explicit zero spacing and unsupported ancestor fields remain rejected. This qualifies the declared defaults and saved XLSX properties; complete appearance remains outside this reference.

`native-exports/pages-simple-v15.4.json` pins the unchanged `nim-iwork/simple.pages` source and its Pages 15.4 DOCX/PDF exports. Reproduce them by opening a copy in Pages and choosing File > Export To > Word or PDF, with password protection off and best PDF image quality. Native controls show Times Regular 12-point text, single spacing, no custom tabs and a 36-point default tab interval. Saved DOCX comparisons check all three paragraphs including the empty paragraph, spacing, paragraph-mark font metrics, page size and margins. The native font label `Times Roman` and decoded `Times-Roman` resolve to Times in this renderer. A one-page native rendering of the saved OfficeIMO DOCX matches all seven word bounds and the 1200-pixel raster after paragraph-mark formatting is preserved. The adjacent manifest records the exact output, fonts and comparison hashes. This qualifies only the recorded fixture and renderer; larger documents, other fonts and producer versions remain open.

The [paragraph-layout creation manifest](native-exports/pages-created-layout-v15.4.json) adds a seven-paragraph document with 1.25/1.5/2 relative spacing, before/after spacing, left/center/right/decimal tabs, TAB text, a blank paragraph and formatting resets. `IWorkPagesParagraphWritingTests.Example()` reproduces its deterministic 8,450-byte input. The 37-object graph uses native paragraph variations, a named default parent and the paragraph table's null end marker. The Apple-saved edit retains the requested layout after close/reopen. Its complete body text, spacing and tabs agree with the native Word export under an independent pinned-schema decoder:

```text
python Build/IWork/verify-pages-creation.py OfficeIMO.TestAssets/Documents/IWorkCorpus/native-exports/pages-created-layout-v15.4.pages --descriptor <pages-descriptor.pb> --out <created-graph.json>
python Build/IWork/extract-pages-native-layout.py OfficeIMO.TestAssets/Documents/IWorkCorpus/native-exports/pages-created-layout-edited-v15.4.pages --descriptor <pages-descriptor.pb> --docx OfficeIMO.TestAssets/Documents/IWorkCorpus/native-exports/pages-created-layout-v15.4.docx --out <native-layout.json>
```

The native PDF and the Pages rendering of OfficeIMO's strict DOCX each contain all 75 words on one page. The DOCX stores the source period separator; Pages 15.4 reimport still shifts `123.45` left by 15.0146 points. The other 74 word bounds agree. The manifest retains both inspected outputs and this rendering boundary. Newly authored fixtures use MIT; the native PDF's installed-font subsets do not establish other font environments.

## Pages paragraph spacing and decimal counters

`native-exports/pages-paragraph-list-v15.4.json` pins a native Pages 15.4 derivative of the licensed `nim-iwork/simple.pages` fixture. Its parent style uses 1.5 spacing and a 180-point left tab; the second paragraph overrides these with 2.0 spacing and a 216-point tab, and the third resets to single spacing and no custom tabs. A non-tiered, level-zero decimal list visibly numbers four paragraphs as 7, 8, 3, 4. The native source, DOCX, complete one-page PDF and screenshot are hashed. Independent numbers-parser 4.19.0 schemas identify storage field `14` starts and zero continuation, the selected decimal kind and disabled tiered numbering.

`native-exports/pages-list-continuation-v15.4.json` pins an Apple Pages 15.4 package produced from an OfficeIMO-owned synthetic DOCX recipe. Its complete one-page native PDF visibly numbers the two levels as 7, 8, 9, 1, 2, 10, 1, 5, 6, 7, 11, 3, 4 across two ordinary-text gaps, child resets and explicit restarts. The source, native DOCX/PDF, independent decode, reconstructed DOCX and both native/composed rendered pages are hashed. Native markers use Calibri 11 points and body text uses ArialMT 12 points, with 1.5 em text offset and an 18-point nested marker indent. Saved DOCX matches the native export’s 330/690-twip left and 330-twip hanging declarations. Pages reimports every counter and exact text/marker starts, but renders the declared 11-point markers at 12 points. The composed PDF retains 11-point marker size and every counter with explicit font substitutions. Font-size roundtrip, broader typography and producer/container coverage remain separate qualification boundaries.

Reproduce the reference by editing a copy in Pages, updating the parent paragraph style, applying the child overrides/resets, and using the numbered-list inspector to set starts at 7 and 3. Save before exporting Word and best-quality PDF with password/comments off. Saved OfficeIMO DOCX checks all paragraph text, effective spacing/tabs, explicit native start overrides and the native 393-twip list indentation. Reader checks all four counters. Original MIT attribution applies; the selected installed Times font and native DOCX marker-font fallback are recorded in the manifest. Other list levels, scripts, mixed marker sizes, interacting indents, text containers and producers need additional native references.

## Native blank-format export

`native-exports/numbers-blank-selectors-v14.5.xlsx` is an Apple Numbers 14.5 export of the unchanged `numbers-parser/cross-table-formulas.numbers` fixture. The adjacent JSON records source/export hashes, export options and ten blank cells in `Main Sheet - Extra Headers`. Reproduce it by opening a copy of the source in Numbers 14.5, choosing File > Export To > Excel, one worksheet per table, and leaving the summary worksheet and password options off. The native export includes a table-title row, so source A8:A17 maps to destination A9:A18. Existing numbers-parser MIT attribution applies.

The nine numeric/text-selector blanks export with built-in Text format 49 (`@`); the isolated percentage-selector blank exports as General. All ten remain blank. This is evidence of this producer's export behavior, not a general selector-precedence rule: the percentage case shows that native export does not retain every authored blank format. OfficeIMO uses the explicit blank selection to assess only the active scalar family; retained inactive catalog entries do not compete with that selection. Native populated-value behavior and rendered appearance remain unqualified.

Native Numbers' Cell Format panel shows source A8 as Automatic and A9 as Text. `numbers-blank-explicit-percentage-v14.5.numbers` is a disposable copy saved after selecting Percentage for A8. The adjacent manifest records the exact before/after storage: selection adds value-storage bit 12 with value `1`, while numeric format key `6` stays unchanged. This demonstrates that a numeric catalog selector alone does not establish active blank formatting. The reader uses this explicit selection for the qualified numeric blank case; the original Automatic blank retains no numeric format. The native Currency → Date and Time → Duration → Automatic sequence and hashed snapshots in `native-exports/numbers-blank-selectors-v14.5.json` qualify blank selection values 2, 3 and 4 alongside numeric 1 and text 5. Switching families retains earlier catalog entries; Automatic clears the selections. These snapshots qualify selection, not broader currency, locale or duration display support.

## Keynote multi-slide editable output

`native-exports/keynote-slide-links-v15.4.json` records a three-slide presentation saved in Keynote 15.4 after importing an OfficeIMO-created PPTX and adding a native frame link. Text links use native slide UUIDs and the frame uses the next-slide action. The source, Apple PDF and Apple PPTX exports are hashed. Independent schema decoding identifies selected slide-node UUIDs in package metadata; pypdf 6.10.0 confirms the native and OfficeIMO PDF targets are 2, 2 and 1. Default editable conversion retains the raw source targets, writes stable PPTX links and rounds the native 3.6-point floating insets with a precision diagnostic. The complete workflow reopens every page with all five text labels and all three links. Inspection of all rendered pages confirms complete visible content with the reported Helvetica font substitution and small baseline differences. Native link activation, exact typography and other slideshow actions remain unqualified.

Reproduce the selected native graph facts in the isolated numbers-parser 4.19.0 evidence environment with the pinned descriptor used by native creation checks:

```text
python Build/IWork/extract-keynote-slide-links.py OfficeIMO.TestAssets/Documents/IWorkCorpus/native-exports/keynote-slide-links-v15.4.key <keynote-descriptor.pb> <selected-links.json>
```

The authored content follows the OfficeIMO MIT license. No standalone Apple fonts are distributed. Native exports use all slides, best PDF quality and no password, presenter notes, comments or build stages.

`native-exports/keynote-simple-v15.4.json` pins the unchanged two-slide Keynote 14.5 source, Keynote 15.4 PPTX/PDF exports and native renderings of Apple and OfficeIMO PPTX files. Default conversion requires complete editable reconstruction and checks slide order, text, notes, selected placeholder geometry, four-point margins, vertical alignment, fit settings and 123% character bullets. Open XML geometry differs by at most two EMUs. Both PPTX renderings have identical rasters at the recorded resolution. The original `.key` PDF has a small ligature/word-advance difference also present in Apple's PPTX round trip. The OfficeIMO import shows an unused Aptos theme-font warning; actual text uses the recorded installed fonts. There is no repair prompt. Exact typography, theme/master preservation and other producer versions remain outside this evidence.

Reproduce selected frame and list declarations with the pinned isolated numbers-parser 4.19.0 environment:

```text
python Build/IWork/extract-keynote-text-layout.py OfficeIMO.TestAssets/Documents/IWorkCorpus OfficeIMO.TestAssets/Documents/IWorkCorpus/native-exports/keynote-text-layout.json --wrapped-dir <probe-output>
```

The optional wrapped probes change only selected body text and keep-lines in a copy of the licensed native fixture. They are test-only derived inputs, not independent producer documents or product writing support. Open working copies in Keynote 15.4 and export PDF with best quality, all slides and optional content/password off; preserve the frozen generated inputs because Keynote can auto-save its working copies. Keep-lines remains true/false after native opening, and both two-slide renders are identical. The strict PPTX retains the same two-line wrapping, with the recorded original-to-PPTX ligature difference. This qualifies only unlinked horizontal single-column fixed frames with the recovered layout. Other pagination contexts retain their strict gates. Existing nim-iwork MIT provenance and attribution apply to the derived probes and exports.

The [native scalar-table creation reference](native-exports/keynote-created-tables-v15.4.json) records a template-free three-slide graph, independent typed-cell and style-network decoding, Keynote 15.4 open/export and edits that survive save/reopen. The fixture includes multilingual text, numbers, booleans, literal equals-sign text and an empty editable table. The [native border reference](native-exports/keynote-table-borders-v15.4.json) adds one header row, one header column, one footer row, hidden region edges and ordered cell overrides. Independent schema decoding, saved/reopened PPTX and inspected all-slide PDF comparisons qualify this border profile. The manifests record locale, font and emoji differences separately.

The [Numbers border reference](native-exports/numbers-table-borders-v14.5.json) contains a controlled four-by-four table imported and edited in Numbers 14.5. Independent protobuf and numbers-parser decoding checks all 64 cell edges against Apple's XLSX export. Saved/reopened OfficeIMO XLSX retains the one- and two-point colors, hidden shared edge, later edit and styled empty cell. Native cell padding requires explicit partial editable acceptance with an omission diagnostic; complete iWork PDF conversion rejects this source. The manifest keeps native PDF appearance and editable edge equivalence separate.

The [Pages border reference](native-exports/pages-table-borders-v15.4.json) checks all 16 values and 64 cell-edge declarations against Pages 15.4 DOCX export. Every physical grid segment has an explicit edit, replacing inactive rounded defaults. Pages opens, saves and reopens the reconstructed DOCX with the edges, ten bold cells and blank D3 intact. The effective dimensions fit DOCX precision; an unused fractional default does not count as precision loss. The [physical row-height reference](native-exports/pages-table-row-height-word-v16.113.3.json) checks the padding-aware fixed-row mapping in native Word 16.113.3 and the shared Word PDF renderer, including vertical merge continuation margins. The [inline-flow reference](native-exports/pages-inline-table-flow-word-v16.113.3.json) qualifies strict DOCX and complete one-page PDF output for this fixed, unmerged grid. Native Word and OfficeIMO retain its 432-by-127-point frame and source position within 0.5 point. Pages DOCX reimport retains content and borders but renders a 150-point frame; font substitution and broader pagination also remain limits. Earlier receipts retain their exact source revisions and acceptance boundaries. Reproduce the source oracle with `Build/IWork/extract-pages-table-borders.py` and destination readback with `Build/IWork/verify-pages-table-borders.py`; each script accepts explicit artifact paths.

## Everyday Apple export references

`native-exports/numbers-formula-formatting-v14.5.json` records a Numbers 14.5 derivative of the licensed cross-table workbook. Formula Tests A9 uses 18-point bold/italic/underline text aligned right; A10 uses 14-point strikethrough text aligned center. The source, XLSX and two-page PDF are hashed. Saved OfficeIMO XLSX checks compare both formula caches, source expressions, font traits, font size, opaque black foreground and alignment with the Apple export after its table-title row is accounted for. The source PostScript name `HelveticaNeue` and exported family name `Helvetica Neue` are recorded aliases for this fixture. Native PDF inspection covers the two source cells; complete appearance and recalculation remain unqualified. Original MIT attribution applies to the derivative.

`native-exports/keynote-image-tracking-v15.4.json` pairs the unchanged licensed image deck with Keynote 15.4 PPTX/PDF exports and an Apple PDF rendered from OfficeIMO PPTX. It qualifies -0.02-em tracking at 116 points as -2.32-point DrawingML spacing, the retained image, and native opening without repair. The recorded render-bound difference and font warning limit the appearance claim.

`native-exports/pages-everyday-v15.4.json` records fresh DOCX/PDF from the unchanged three-page Pages fixture. Executable checks compare meaningful body order, section header/footer text, image count and selected cells across all three tables. Seven date/currency cells have unqualified locale display. Selected typography and native pagination remain bounded by the explicit partial policy.

`native-exports/numbers-cross-table-v14.5.json` records XLSX/PDF exports and the scratch copy auto-saved by Numbers. The original upstream fixture remains unchanged. Checks compare 14 worksheet names and all 529 native populated values with caches, including 109 formula caches, after accounting for Apple's added title row. One formula has no native cached error value. Formula rich text, ordering, recalculation, cache freshness and appearance remain outside this comparison. Existing upstream MIT notices apply; manifests record source/artifact hashes, versions, export settings and installed-font provenance. No standalone fonts are distributed.

## Pages editable body output

The unchanged Pages 14.1 `iwork-converter/a.pages` fixture exercises a 45-paragraph source body through explicitly accepted partial DOCX conversion, save and reopen. The corpus test compares non-empty body paragraphs in order and checks the title's centered alignment, bold styling, 26-point size and declared `SimSun` font family. It validates the saved Open XML document. The fixture declares the font name; this test does not distribute or render that font. Inline objects remain incomplete, and native Apple DOCX export, page count and rendered equivalence remain unqualified.

`pages-section-selection.json` records the two native sections' header/footer selection flags and template references. Both store first/even/default templates while selecting the default for all pages. Field names come from the pinned independent TP schema recorded in the manifest. Reproduce the declarations with `python Build/IWork/extract-pages-section-selection.py OfficeIMO.TestAssets/Documents/IWorkCorpus output.json` in the opt-in numbers-parser 4.19.0 environment. Enabled alternate templates and first-page suppression have synthetic saved-DOCX coverage; native export qualification for those settings remains open.

`pages-word-rendering.json` records a bounded manual inspection of this conversion and `picodocs/sample-v14.4.pages` in Microsoft Word for Mac 16.112.4. Both DOCX files opened without repair prompts; representative titles, body text, a table and the embedded chart rendered visibly. The recorded baseline first-page comparison shows ordered items rendered as bullets, omitted header/footer content and different pagination. The numbering-remediation entry confirms the regenerated decimal list visibly renders as 1., 2., 3. in the same Word version. The header/footer-remediation entry confirms that the default header and footer visibly render on the first page and the header remains visible on page two. This qualifies destination opening and the inspected content only; native Apple export and rendered equivalence remain unqualified.

## Pages paragraph layout declarations

`pages-paragraph-layout.json` records line-spacing and custom-tab messages in selected body paragraph styles and their parent chain in the unchanged `picodocs/sample-v14.4.pages`. The pinned independent parser identifies twelve declarations across seven styles, including empty messages whose single-line/empty-custom-tab defaults are qualified by the paired Numbers 14.5 evidence below. The manifest also exposes independently decoded relative multipliers and tab positions/alignments without leaders. The reader retains those explicit values and source paths for remaining unqualified layout. DOCX/PPTX spacing properties are checked separately from rendered appearance. This evidence does not qualify Apple exports or rendered equivalence.

```sh
python Build/IWork/extract-pages-paragraph-layout.py OfficeIMO.TestAssets/Documents/IWorkCorpus OfficeIMO.TestAssets/Documents/IWorkCorpus/pages-paragraph-layout.json
```

## Pages list-style declarations

`pages-list-styles.json` records the selected body list styles and paragraph-data boundaries in the unchanged `picodocs/sample-v14.4.pages` fixture. The ordered span selects native number labels and decimal numbering (`number_types = 0`), with no string labels and tiered numbering disabled. The preceding bullet span selects string labels and the `Symbol` marker font. The reader retains inherited marker fonts and explicit font clearing; DOCX numbering and PPTX bullets preserve explicit fonts independently of run fonts. Saved DOCX coverage checks the native bullet font, and saved PPTX coverage checks a distinct explicit font. Font installation and rendered glyph equivalence remain separate qualification requirements. The independent schema enum names identify native marker kinds separately: none (`0`), image (`1`), string (`2`) and number (`3`). Literal string markers never imply a numbering format. Image markers retain incomplete-formatting and source-declaration diagnostics instead of using dormant string labels. The reader projects initial markers for decimal, upper/lower Roman and upper/lower alphabetic numbering, each with a dot, paired parentheses or a closing parenthesis. Selected native numbering kinds take precedence over dormant string labels. Paragraph-data field `6/1/2` supplies explicit list levels, as identified by the pinned independent parser source in the manifest. Explicit levels take precedence over indentation and remain distinct in the style cache; absent or rejected values use the bounded indentation fallback with rejected metadata reported. Unknown values or an uncorrelated restart flag in `6/1/3` retain source-declaration evidence and incomplete-formatting diagnostics. Saved DOCX regression coverage checks the fixture’s three decimal items share one numbering instance; saved DOCX/PPTX cases check all fifteen qualified initial numbering formats. The separate Pages 15.4 paragraph/list reference qualifies level-zero decimal starts/restarts. The [continuation reference](native-exports/pages-list-continuation-v15.4.json) extends the non-tiered decimal contract to levels 0 and 1, parent resets and ordinary-text gaps. Its observed `6/1/3 = 1` restart flag requires a positive start at the same paragraph boundary. This older fixture retains unqualified metadata and still requires partial editable reconstruction. Deeper levels, numbering scripts, tiered numbering and continuation through other numbered or bullet styles remain unassessed.

Reproduce with the pinned opt-in numbers-parser 4.19.0 provider:

```bash
python Build/IWork/extract-pages-list-styles.py OfficeIMO.TestAssets/Documents/IWorkCorpus OfficeIMO.TestAssets/Documents/IWorkCorpus/pages-list-styles.json
```

## Pages inline attachment positions

`pages-inline-anchors.json` records five image/table attachment positions from the unmodified `iwork-converter/a.pages` and `picodocs/sample-v14.4.pages` fixtures. The existing provenance, hashes and license notices above apply. Positions are UTF-16 offsets in native text storage, before object-marker removal. The manifest is extracted through the independent numbers-parser 4.19.0 schemas:

```bash
python Build/IWork/extract-pages-inline-anchors.py OfficeIMO.TestAssets/Documents/IWorkCorpus OfficeIMO.TestAssets/Documents/IWorkCorpus/pages-inline-anchors.json
```

OfficeIMO compares native attachment/drawable identities and offsets with this evidence. Saved/reopened DOCX and Reader tests check the representative fixture's object order. This qualifies source attachment decoding and destination placement; native Apple export, pagination, wrapping and rendered equivalence remain unqualified.

## Pages automatic row sizing

`pages-table-sizing.json` records the native automatic-resize setting and individual row heights of all three tables in the unchanged `picodocs/sample-v14.4.pages` fixture. It uses the same pinned source, hash and MIT provenance above. Reproduce the manifest with independent numbers-parser 4.19.0 schemas:

```bash
python Build/IWork/extract-pages-table-sizing.py OfficeIMO.TestAssets/Documents/IWorkCorpus OfficeIMO.TestAssets/Documents/IWorkCorpus/pages-table-sizing.json
```

Saved/reopened DOCX retains the declared 16.5-point heights as minimum constraints, so wrapped cell content can increase row height. The embedded package preview provides a bounded visual counterexample to treating these rows as fixed. This qualifies the source setting and destination constraint; native Apple export, complete styling and pagination remain unqualified.

## Pages selected cell fills and layout

`pages-cell-fills.json` records 64 selected modern cell-style keys, parent chains, solid/no-fill declarations, four-sided padding and vertical alignment across the same three native Pages tables. It includes four empty cells whose explicit fills must survive sparse projection. Reproduce it with the pinned independent numbers-parser 4.19.0 schemas:

```bash
python Build/IWork/extract-pages-cell-fills.py OfficeIMO.TestAssets/Documents/IWorkCorpus OfficeIMO.TestAssets/Documents/IWorkCorpus/pages-cell-fills.json
```

Saved/reopened DOCX compares each selected fill, padding side and vertical alignment with this manifest. All 64 selected cells have middle alignment and four-point padding. Empty fill declarations clear parent colors; table-role defaults and banding are outside this evidence. Native Apple export and complete appearance remain unqualified.

## Cross-table rectangular formulas

`numbers-parser/cross-table-formulas.numbers` is the unmodified upstream `tests/data/create-formulas.numbers` at revision `1c6c5c3d2e29a9abb601596678089f0a6c85d64c`, covered by the existing numbers-parser MIT notice. Its adjacent JSON manifest records sixteen `COUNTA` formulas with cross-table rectangular references, all combinations of absolute and relative endpoint coordinates, native target UUIDs and cached numeric values. `Build/IWork/extract-cross-table-formula-fixture.py` reads the pinned package through numbers-parser 4.19.0 to reproduce the manifest without rewriting the source:

```bash
python Build/IWork/extract-cross-table-formula-fixture.py OfficeIMO.TestAssets/Documents/IWorkCorpus/numbers-parser/cross-table-formulas.numbers OfficeIMO.TestAssets/Documents/IWorkCorpus/numbers-parser/cross-table-formulas.json --whole-axis-output OfficeIMO.TestAssets/Documents/IWorkCorpus/numbers-parser/whole-axis-formulas.json
```

OfficeIMO tests compare all sixteen reconstructed source ranges and saved/reopened XLSX expressions with the manifest, preserving mixed endpoint flags and numeric caches. Separate boundary tests cover forward targets, normalized colliding names, and missing, inactive or ambiguous identities; unresolved targets never become complete local formulas. This qualifies the finite rectangular subset through independent reference and cache evidence.

`whole-axis-formulas.json` uses the same hashed native package for 59 row/column references, including coordinate-backed header-name aliases. The extractor resolves axes and target header/footer metadata through numbers-parser, computes `SUM`, `COUNT` or `COUNTA` over the selected current body, and verifies every result against the native cache. Saved-output tests compare exact fixed body bounds, endpoint flags, caches and the typed approximation diagnostic. Synthetic cases cover footers, normalized colliding names, ambiguous metadata and empty bodies. Native labels and automatic expansion are not preserved in fixed XLSX ranges. Native Apple export, appearance, recalculation and cache freshness remain unqualified.

## Formula function identities

`numbers-parser/function-identities.json` records 75 native function identifiers from the pinned numbers-parser 4.19.0 function map, plus the expressions, function nodes and typed caches of three cases in the unmodified `cross-table-formulas.numbers` fixture. OfficeIMO checks every selected identity, rejects unknown identifiers and invalid argument counts, and saves/reopens the native nested `OR`, `POWER`, and cross-table `TEXTJOIN` cases. The `TEXTJOIN` descriptor records independently resolved coordinates and a join of current referenced values that agrees with its cache. Its saved-output test checks the XLSX compatibility prefix, typed string cache, and recalculation after changing a referenced value and testing both empty-cell policies. The additional `MINA`, `MAXA` and `AVERAGEA` identities use the independent map; synthetic native-format cases check saved output separately from the three unmodified document cases. Sample argument counts exercise reconstruction; they do not qualify evaluation for arbitrary argument types. Reproduce the evidence without rewriting the package:

```bash
python Build/IWork/extract-function-identities.py OfficeIMO.TestAssets/Documents/IWorkCorpus/numbers-parser/cross-table-formulas.numbers OfficeIMO.TestAssets/Documents/IWorkCorpus/numbers-parser/function-identities.json
```

The existing numbers-parser MIT notice covers the provider and fixture. The manifest does not qualify Apple export, appearance, recalculation or cache freshness.

## Empty hidden-state extents and disabled filters

`empty-hidden-states.json` describes five table models in the unchanged `nim-iwork/simple.numbers`, `picodocs/sample-v14.4.pages` and `keynotekit/tabledeck-v15.2.1.key` packages. Their pinned provenance and license notices above apply. Independent numbers-parser 4.19.0 schemas verify empty base/summary hidden states, correct axis directions, zero hidden counts and ten disabled filter sets without rules. Reproduce the manifest without rewriting the packages:

```bash
python Build/IWork/extract-empty-hidden-states.py OfficeIMO.TestAssets/Documents/IWorkCorpus OfficeIMO.TestAssets/Documents/IWorkCorpus/empty-hidden-states.json
```

OfficeIMO checks the source hashes, selected model identities and absence of new visibility warnings or declaration failures. This qualifies the empty/disabled path. The base user-hidden column qualification below supplies independent positive selection evidence; active filtering, collapsed groups and Apple export equivalence remain unqualified.

## Individual table dimensions

`numbers-parser/individual-dimensions.numbers` is produced by numbers-parser 4.19.0 using `Build/IWork/create-dimension-fixture.py`. It declares three row heights (20, 10, and 30 points) and two column widths (40 and 20 points), with source labels in the first column. The generator reopens the package through the independent producer; OfficeIMO tests read the declared dimensions and save/reopen the XLSX result. This fixture also exercises the valid seven-field tile envelope with the wide-row flag. It qualifies declared dimensions, not Apple automatic row sizing or visual equivalence.

To recreate the semantic fixture in an isolated Python 3.10+ environment, install `numbers-parser==4.19.0`, then run:

```bash
python Build/IWork/create-dimension-fixture.py OfficeIMO.TestAssets/Documents/IWorkCorpus/numbers-parser/individual-dimensions.numbers
```

The producer creates new package identifiers and timestamps, so regenerated bytes can differ. The committed fixture has the checksum recorded below. The existing numbers-parser MIT notice covers its bundled template and implementation.

## Numeric formats

`numbers-parser/number-formats.numbers` is produced by numbers-parser 4.19.0 using `Build/IWork/create-number-format-fixture.py`. Its adjacent generated JSON manifest records the package hash, thirteen numeric values, exact source coefficient/exponent text, selected format metadata and independent producer display strings. Cases cover decimal precision, automatic formatting, grouping, all four negative styles, percentage scaling and zero. The generator reopens the package through the producer before writing the oracle.

```bash
python Build/IWork/create-number-format-fixture.py OfficeIMO.TestAssets/Documents/IWorkCorpus/numbers-parser/number-formats.numbers
```

OfficeIMO tests compare metadata, save/reopen XLSX numeric values and compare image-snapshot text with the producer. The fixture also exposes finite Decimal128 values above fifteen significant digits; their exact source text and explicit approximation survive reconstruction. This evidence does not qualify Apple native rendering or full style fidelity. Regenerated package identifiers and timestamps can change the bytes; compare the semantic manifest cases when reproducing the fixture.

## Currency formats

`numbers-parser/currency-formats.numbers` is produced by numbers-parser 4.19.0 using `Build/IWork/create-currency-format-fixture.py`. The adjacent generated manifest records eleven currency cases across six identifiers, precision, grouping, negative styles, accounting settings, and independent source display strings. Destination display strings specify the portable identifier-prefix contract, including its explicit approximation report; they are not Apple appearance oracles.

```bash
python Build/IWork/create-currency-format-fixture.py OfficeIMO.TestAssets/Documents/IWorkCorpus/numbers-parser/currency-formats.numbers
```

The generator saves and reopens the package before recording metadata and source display text. OfficeIMO checks source semantics and saved/reopened XLSX numeric values, image text, and red styles. The existing numbers-parser MIT notice covers the template and implementation. Regenerated package identifiers and timestamps can change bytes; compare the semantic manifest cases.

## Scientific formats

`numbers-parser/scientific-formats.numbers` is produced by numbers-parser 4.19.0 using `Build/IWork/create-scientific-format-fixture.py`. Its generated manifest records fourteen scientific selections, exact source numeric text, producer-decoded values, correctly rounded portable values, and selected mantissa precision. Twelve explicit-precision display strings come from the reopened producer. Its automatic sentinel is rendered as 253 fractional places, so the two automatic cases qualify metadata only; destination expectations describe OfficeIMO's approximate display contract.

```bash
python Build/IWork/create-scientific-format-fixture.py OfficeIMO.TestAssets/Documents/IWorkCorpus/numbers-parser/scientific-formats.numbers
```

The scientific and numeric-format generators share `fixture_number_values.py` to derive portable numeric values and precision evidence from the stored finite Decimal128 coefficient and exponent. OfficeIMO tests compare source metadata and exact numeric values, then save/reopen XLSX and check display text. This is independent-producer evidence, without an Apple native export or appearance claim. Regenerated identifiers and timestamps can change package bytes; compare semantic manifest cases. The existing numbers-parser MIT notice covers the template and implementation.

## Fraction formats

`numbers-parser/fraction-formats.numbers` is produced by numbers-parser 4.19.0 using `Build/IWork/create-fraction-format-fixture.py`. Its generated manifest records twenty-six cases across all nine denominator modes, stored numeric precision, source metadata, and reopened producer display strings. Eighteen display examples qualify the independent oracle. Eight destination alternatives cover negative whole parts, rollover, and midpoint ties where the producer's output loses information, remains unnormalized, or differs in rounding policy. Those alternatives qualify OfficeIMO's reported display approximation, without claiming producer or Apple appearance equivalence.

```bash
python Build/IWork/create-fraction-format-fixture.py OfficeIMO.TestAssets/Documents/IWorkCorpus/numbers-parser/fraction-formats.numbers
```

The generator reuses `fixture_number_values.py` for exact Decimal128 text and portable numeric values. The near-quarter-midpoint example retains the value actually stored by the writer, which lies below the midpoint. Tests compare saved/reopened XLSX numeric types, values, format codes, and image-snapshot text. Regenerated identifiers and timestamps can change bytes; compare semantic manifest cases. The existing numbers-parser MIT notice covers the template and implementation.

## Independent Apple exports

`native-exports/numbers-formulas-v14.5.json` records exports of `numbers-parser/test-10-formulas.numbers` made with Apple Numbers 14.5 (build 7045.0.17). The unmodified XLSX and PDF references include artifact hashes, export settings, source licensing, font provenance, and qualification limits. Numbers exports one worksheet per table and inserts a title row; the manifest accounts for that row when comparing source coordinates and formulas.

`IWorkAppleExportQualificationTests` saves and reopens OfficeIMO's XLSX, compares all 28 formula expressions with the independent export, and compares 24 stable typed cached values. Two volatile `NOW` caches are excluded because Apple recalculates them; two error formulas have no native cached value. Generic source error text is retained with an approximation diagnostic and is not qualified as a native Excel error value. The PDF supplies a two-page visual reference; rendered equivalence is not yet qualified. Its axis-aligned table grids have 98-point columns and 20.07-point rows. A separate saved-output check compares declared column widths and bounds the difference between declared row heights, Apple XLSX heights and the measured PDF grid; it does not implement Apple automatic row sizing.

The `pdf.tableGeometry` manifest field is generated from the hashed PDF by the opt-in `pdfplumber` tool (version 0.11.9 for this evidence):

```bash
python Build/IWork/update-numbers-export-geometry.py OfficeIMO.TestAssets/Documents/IWorkCorpus/native-exports/numbers-formulas-v14.5.json
```

The extractor accepts this pinned two-table fixture only, verifies the PDF hash and checks grid line counts against its source coordinate mapping before updating the manifest. It is not part of normal restore or runtime dependencies.

## Fixture checksums

| Fixture | SHA-256 |
|---|---|
| `iwork-converter/a.key` | `929347827a7478c123dd3e3828e9751b5cf2ae977d2edd2a5f0774014fc4fefc` |
| `iwork-converter/a.numbers` | `43afd5a01cff9283fea5a11716d96f696b5853b75f82a3fe4469f240dfb82947` |
| `iwork-converter/a.pages` | `8481e3071c8ea1cc9543354bcd1ff66f79e6c2c8686096fa3ae60ab96fd49ebf` |
| `keynotekit/imagedeck-v15.2.1.key` | `a9af589197588e04ee52388b0aa6c2dad1110e5d6db814b58afe543831cf2128` |
| `keynotekit/tabledeck-v15.2.1.key` | `384962b1fff18abc5a901b59dc5f8820c2a959977f18f90dc9cd10095bdd0a56` |
| `nim-iwork/simple.key` | `ba95755df82ceb0ca834e1e03e2777c34fad906320d8336b4f3fefc6b48607eb` |
| `nim-iwork/simple.numbers` | `d0b00d9cae5985cccaa3b2fb251fae92eb0e38360fb4b5df8b4350eb658f752b` |
| `nim-iwork/simple.pages` | `5aee6d03277d2db2104f593e64afe081dec539f0117b97124b6f99158124c93e` |
| `numbers-parser/individual-dimensions.numbers` | `ccecc5494a71b9943e9a37641f843d0e8ce4e7c3531d72a50031dd7c23b38762` |
| `numbers-parser/currency-formats.numbers` | `6f45fa942ab26b2e9a3c2487fec491d5c43e85cfe1c191520134d5b64ac61475` |
| `numbers-parser/number-formats.numbers` | `1fb277a6897c0fc4387cd50a51c57de0a08323b694b98a5926b9bc5190717b1f` |
| `numbers-parser/scientific-formats.numbers` | `23bfaa5cf394c46ec5192e9612aca75c29f00e84331f11ffc338ef338aa5847d` |
| `numbers-parser/fraction-formats.numbers` | `da140e7eeac3122505690af655a5896aa056573dd15826a2f84d035388fbff55` |
| `numbers-parser/issue-102-v15.1.numbers` | `88a9fa7be095d03004478393a87a4a97602d7468f839d067ec9118c524c55176` |
| `numbers-parser/cross-table-formulas.numbers` | `9371c5b1d6ee4dfa17569097f064eba9c67f804d88b48638efbbeeb459d07dd4` |
| `numbers-parser/test-10-formulas.numbers` | `dd85bad68898ce5b065f277c0b9be1f3c32d696e3baa6b09d3614bbd35a5249f` |
| `numbers-parser/test-9-merges.numbers` | `d640c0012d629834161827cb2f564d0966d24e159586f82426a69c12a8f334cf` |
| `picodocs/sample-v14.4.pages` | `4714477138d0a4090fc2ee2ba2ebb6adcd0fb6ce20a28897a6247a8e17d1ddce` |

## Keynote unbanded table-fill defaults

`keynote-table-fill-defaults.json` records the unchanged Keynote 15.2.1 table fixture's disabled banding and explicit no-fill role declarations through pinned independent numbers-parser 4.19.0 schemas. Reproduce it with:

```sh
python Build/IWork/extract-keynote-table-fill-defaults.py OfficeIMO.TestAssets/Documents/IWorkCorpus OfficeIMO.TestAssets/Documents/IWorkCorpus/keynote-table-fill-defaults.json
```

All nine saved/reopened PPTX cells preserve explicit no-fill and suppress destination theme backgrounds. This qualifies source declarations and saved properties; role intersections, banded appearance, Apple exports and complete styling remain unqualified.

## Native Numbers table banding

`native-exports/numbers-banding-v14.5.numbers` is an OfficeIMO-authored fixture saved by Numbers 14.5, covered by the repository MIT license. Three eight-row tables vary the header-row count from zero to two and retain one header column and one footer row. Row seven's “Footer” label is body content; the actual footer is the empty eighth row. No selected cell-fill overrides are stored.

The matching Apple XLSX and three-page PDF exports qualify the alternating body-row pattern and region intersections. The XLSX inserts a table-title row. The manifest records all 72 native position fills, source and export hashes, producer version, export settings and limits. Reproduce the source/schema and XLSX oracle extraction with pinned opt-in numbers-parser 4.19.0:

```sh
python Build/IWork/extract-numbers-banding.py OfficeIMO.TestAssets/Documents/IWorkCorpus OfficeIMO.TestAssets/Documents/IWorkCorpus/native-exports/numbers-banding-v14.5.json
```

This evidence covers opaque sRGB region and band fills in this fixture. It does not establish complete appearance, selected override behavior in Apple exports, other producer versions, or Pages and Keynote banded exports.

## Native Numbers function exports

`native-exports/numbers-functions-v14.5.json` records two OfficeIMO-authored Numbers 14.5 fixtures and their Apple XLSX exports. The matching intake workbooks contain cache-free OOXML formulas built with Python standard-library ZIP/XML, independently of OfficeIMO. Numbers imports, evaluates and writes the native files and exports. All assets are covered by the repository MIT license.

The manifest records 26 expressions, native function IDs and argument counts, numeric cache comparisons, producer version and artifact hashes. It exercises `OFFSET` as a scalar and as a `SUM` range argument, `PROB` exact and interval bounds, zero probabilities and invalid distributions, and fixed and variable `RANDBETWEEN` bounds. Reproduce the manifest with pinned opt-in numbers-parser 4.19.0:

```sh
python Build/IWork/extract-numbers-functions.py OfficeIMO.TestAssets/Documents/IWorkCorpus OfficeIMO.TestAssets/Documents/IWorkCorpus/native-exports/numbers-functions-v14.5.json
```

The fractional random cases returned zero in this native import and remain outside local evaluation qualification. Random caches are snapshots. Native error cells have no XLSX error cache, so this evidence does not qualify error-code equivalence. Saved/reopened conversions retain the expressions and diagnosed caches; fresh shared-owner evaluation and edits to referenced inputs are tested separately. The PDF is a native reference, not a qualification of complete layout, implicit intersection, dynamic spills or other producer versions.

## Fractional random bounds

`native-exports/numbers-fractional-bounds-v14.5.json` records 16 cache-free expressions imported, evaluated and saved by Numbers 14.5, with its XLSX export and a separate native Excel 16.113.3 recalculation. The intake is OfficeIMO-authored; the results come from the native applications. Artifact hashes, function nodes, typed caches, screenshots and producer versions are retained. Reproduce the manifest with numbers-parser 4.19.0:

```sh
python Build/IWork/extract-numbers-functions.py OfficeIMO.TestAssets/Documents/IWorkCorpus OfficeIMO.TestAssets/Documents/IWorkCorpus/native-exports/numbers-fractional-bounds-v14.5.json --profile fractional-random
```

Seven nonempty integer intervals agree across the native applications and fresh OfficeIMO evaluation. Empty integer intervals expose a compatibility limit: Numbers returns zero, while Excel can return the rounded lower bound. OfficeIMO retains those formulas and caches without local evaluation. Native Numbers errors for reversed bounds have no exported error cache; fresh Excel evaluation returns `#NUM!`. Random sequences, full appearance and additional producer versions remain unqualified.

## Base user-hidden columns

`numbers-parser/user-hidden-columns.json` describes six hidden columns across three tables in the unchanged `numbers-parser/cross-table-formulas.numbers` package. Its pinned upstream revision and MIT notice above apply. Build metadata records Numbers `M14.3-7042.0.76-4` and the Blank 11.2 template. Independent numbers-parser 4.19.0 schemas resolve selected UUIDs through nonidentity column permutations and verify their inverse vectors, zero legacy hidden counts and empty hidden cells. Reproduce the manifest:

```bash
python Build/IWork/extract-user-hidden-columns.py OfficeIMO.TestAssets/Documents/IWorkCorpus/numbers-parser/cross-table-formulas.numbers OfficeIMO.TestAssets/Documents/IWorkCorpus/numbers-parser/user-hidden-columns.json
```

OfficeIMO checks the source checksum, recovered positions and every column’s saved XLSX hidden attribute. Synthetic format-boundary cases cover positive rows, populated hidden cells, Reader inclusion and DOCX/PPTX partial policy. This qualifies base user-hidden column decoding and XLSX metadata; Apple export/render equivalence, active filters, pivot hiding, summary states and collapsed groups remain open.

## Single-cell and endpoint formulas

`numbers-parser/single-cell-formulas.numbers` and `endpoint-formulas.numbers` are unchanged upstream `tests/data/test-all-formulas.numbers` and `test-extra-formulas.numbers` at revision `d3836ebda1110b5c13b8722642ca61111fe8e865`. The existing numbers-parser MIT notice applies. The first package records an XLSX import followed by Numbers 11.1 through 13.1 saves; the second records CSV import and Numbers 11.1/12.2 saves. Their adjacent manifests retain source checksums and build metadata. Reproduce each with independent numbers-parser 4.19.0:

The `single-cell-formulas.json` reference evidence also covers `ROW`, `ROWS`, and `COLUMNS`: five coordinate-backed expressions and five current-cell/scalar expressions in the unchanged Numbers source. The extractor computes row positions and rectangular dimensions independently; native array operands are outside this subset. The manifest records exact source coordinates and cached values. Destination recalculation is tested separately with synthetic native-format input because the original workbook retains its diagnosed pre-1900-date conversion fallback.


```bash
python Build/IWork/extract-single-cell-formulas.py OfficeIMO.TestAssets/Documents/IWorkCorpus/numbers-parser/single-cell-formulas.numbers OfficeIMO.TestAssets/Documents/IWorkCorpus/numbers-parser/single-cell-formulas.json
python Build/IWork/extract-single-cell-formulas.py OfficeIMO.TestAssets/Documents/IWorkCorpus/numbers-parser/endpoint-formulas.numbers OfficeIMO.TestAssets/Documents/IWorkCorpus/numbers-parser/endpoint-formulas.json
```

The native `COMBIN(8,2)` case retains the producer result of 28 and agrees with independent integer-combination computation. Saved/reopened conversion and recalculation are qualified separately with synthetic native-format input; the original workbook keeps its diagnosed pre-1900-date fallback.

Six `FACT` expressions cover zero, one, fractional truncation, five, 150 and an addition to the large result. Independent integer-factorial computations agree with the retained producer caches within the recorded relative tolerance of 1e−14. The separate native negative-input error case records its error classification without asserting a specific XLSX error code. Destination negative-input and overflow behavior is qualified separately through the shared Excel evaluator.

Ten additional native numeric cases cover `EVEN` and `ODD`, including signed fractions, zero and inputs already at the requested parity. The extractor uses independent integer arithmetic to verify their cached values. Synthetic native-format conversions test the destination evaluator separately, preserving the original workbook's pre-1900-date safety boundary.

The endpoint manifest also records `DEGREES(PI()÷2)` and `IFERROR(1÷0,"DIV ZERO")`, with explicit XLSX formula spellings that retain their original native text. Independent computations agree with the numeric/text caches. Both unchanged native expressions are checked through editable conversion, saved/reopened XLSX and local recalculation.

The first manifest also records ten `ISBLANK`, `ISERROR`, `ISNUMBER`, `ISTEXT` and `CONCATENATE` cases. Independent computations distinguish blank references from empty text and numeric literals from text, identify division by zero, and join ASCII literals. Their native caches are Boolean or text. The manifest gives the division expression an explicit XLSX spelling (`1/0`) while retaining its native formula (`1÷0`). Saved/reopened synthetic conversions separately test stale-cache preservation and local recalculation after cross-table type edits.

The manifests resolve node-36 identities, signed relative offsets and mixed absolute flags through independent native schemas. Independent computations agree with the qualified reference caches. The same extractor records scalar expressions and native function nodes: `EXACT` (matching and case-different strings), `LOWER` and `TRIM` in the first package, and `NOT` plus `UPPER` in the second. Independent numeric/ASCII computations agree with their Boolean/text caches. OfficeIMO checks the first package’s reconstructed reference and scalar expressions with their typed caches. Its pre-1900 dates still require whole-workbook XLSX fallback; the fixture is not edited to bypass that safety rule. The first manifest also records 67 numeric expressions: nine `INT`, `MOD` and `SQRT` cases plus sixteen `SIGN`, `TRUNC`, `ROUNDUP` and `ROUNDDOWN` cases with negative literals and nested `ABS`; independent numeric computations agree with their reconstructed source formulas and numeric caches. Independent decimal computations cover signed digit arguments, omitted `TRUNC` digits and positive/negative/zero signs. Equivalent exponent and decimal literal spellings are normalized when comparing reconstructed formulas. Eleven additional `EXP`, `LN`, `LOG` and `LOG10` expressions cover nesting, exponentiation and omitted/explicit bases. Their exact producer caches agree with independent transcendental computations within the manifest’s recorded relative tolerance of 1e−14 and absolute tolerance of 1e−15; this does not replace the source cache with a freshly calculated value. Synthetic native-format inputs separately qualify editable XLSX caches and local recalculation for these thirteen functions, including explicit `TRUNC` digits. Ten `CEILING`/`FLOOR` expressions also qualify signed and decimal factors, zero significance and the native omitted node-22 operand in `FLOOR(0,)`. Independent decimal computations agree exactly with stored Decimal128 values; raw bytes and provider floats remain separate evidence. The stored `0.24` is reported as `0.24000000000000002` by the independent provider’s binary scaling. The second package qualifies saved/reopened `COUNTBLANK`, `MAX`, single-cell `OFFSET`, `NOT` and `UPPER` expressions with numeric/Boolean/text caches. Recalculation follows edited references and logical/text literals. Synthetic packages cover all absolute/relative combinations, forward and normalized names, unresolved identities, malformed coordinates, endpoint whitespace and incompatible endpoint targets. Apple export, rendering and broader native recalculation remain unqualified.

## Native root cell comments

`cell-comments/native-roots.numbers` is the unchanged MIT-licensed `fixtures/olekristensen-v26.3-demo06-formulas-round8.numbers` from [cupertino-files](https://github.com/den-frie-vilje/cupertino-files) at revision `6879e6ed49e0eed7e7f393ff4ee558dc2dbec561`. The upstream attribution records a macOS Numbers save. The adjacent MIT notice retains copyright 2026 Ole Kristensen. The fixture contains the author's published test notes, display name and native author identifiers; it is retained deliberately as native comment evidence.

With independent numbers-parser 4.19.0 installed, reproduce its hash-pinned manifest:

```sh
python Build/IWork/extract-cell-comments.py OfficeIMO.TestAssets/Documents/IWorkCorpus/cell-comments/native-roots.numbers OfficeIMO.TestAssets/Documents/IWorkCorpus/cell-comments/native-roots.json
```

The three selected roots qualify exact text, display author, timestamps, source record identities and saved/reopened XLSX cell anchors. Synthetic packages cover empty commented cells in all three formats, unresolved references, duplicate keys/text, unsupported replies and text/catalog limits. The fixture contains no replies and is not an Apple export or rendered-appearance oracle.

## Table catalog type and kind evidence

`table-catalog-contract.json` records independent registry aliases, source hashes and declared catalog links across the checked-in native-format corpus. Reproduce it with opt-in numbers-parser 4.19.0:

```sh
python Build/IWork/extract-table-catalog-contract.py OfficeIMO.TestAssets/Documents/IWorkCorpus OfficeIMO.TestAssets/Documents/IWorkCorpus/table-catalog-contract.json
```

The 312 links across 24 packages all use type `6005` and explicitly declare the kind expected by the owning store field. The independent registry also maps `6201` to `TST.TableDataList`; synthetic selected-value cases exercise that alias for strings, formulas, rich text, styles, comments and number formats. There is no native `6201` sample in this corpus. Model records may be inactive, and the inventory does not qualify selection, Apple exports or rendered appearance.

## Selected scalar format declarations

`numbers-parser/cell-format-selectors.json` records default text/Boolean and selected date/time/duration format metadata from the unchanged `cross-table-formulas.numbers`, `endpoint-formulas.numbers`, `single-cell-formulas.numbers` and `issue-102-v15.1.numbers` packages. Their source hashes and existing provenance/licensing above apply. The extractor retains 15 distinct representative declarations across the packages, their exact serialized format messages and counts of selected format flags. Default text type `260` and Boolean type `1` have no additional fields; date/time type `261` retains patterns and duration type `268` retains unit/style settings. OfficeIMO checks all date/duration selector counts and the representative feature classifications. This qualifies source metadata and default-format acceptance, without Apple export or appearance claims.

Reproduce with opt-in numbers-parser 4.19.0:

```sh
python Build/IWork/extract-cell-format-selectors.py OfficeIMO.TestAssets/Documents/IWorkCorpus/numbers-parser OfficeIMO.TestAssets/Documents/IWorkCorpus/numbers-parser/cell-format-selectors.json
```


`numbers-parser/blank-number-formats.json` records 35 blank cells with numeric selectors in the same unchanged corpus, including the explicit format-selection field. Twenty-five cells select numeric formatting with value `1`: four integer-format cells in `single-cell-formulas.numbers` and 21 grouped-number cells in `issue-102-v15.1.numbers`. Tests compare their projected formats with the independent metadata and verify the 21 Numbers 15.1 cells remain blank with their formats after XLSX save/reopen. Conversion explicitly accepts other partial reconstruction in that workbook; this is not whole-document fidelity qualification. The remaining isolated percentage catalog entry has no explicit selection and stays inactive, as confirmed by the [native Automatic/Percentage control](#native-blank-format-export). Nine cells explicitly select the default text format with value `5`; tests preserve `IWorkNumberFormatKind.Text` without activating their dormant numeric entries and compare saved/reopened XLSX Text formats (`@`) with the native export. The bounded native export records Text for those nine cells.

```sh
python Build/IWork/extract-blank-number-formats.py OfficeIMO.TestAssets/Documents/IWorkCorpus/numbers-parser OfficeIMO.TestAssets/Documents/IWorkCorpus/numbers-parser/blank-number-formats.json
```


`numbers-parser/duration-format.json` qualifies the fixed abbreviated hour/minute selection in source Table 1 C5 of unchanged `test-10-formulas.numbers`. Its source settings, seconds and independent display text agree with the paired Numbers 14.5 XLSX C6 cache and `[h]"h" m"m"` format after the native title-row offset. Other units, styles, automatic selection, negative rounding and locale or appearance equivalence remain unqualified. Existing source/export hashes, licensing and producer provenance apply. Regenerate with the pinned opt-in provider:

```sh
python Build/IWork/extract-duration-format.py OfficeIMO.TestAssets/Documents/IWorkCorpus OfficeIMO.TestAssets/Documents/IWorkCorpus/numbers-parser/duration-format.json
```

`numbers-parser/date-formats.json` is generated by `Build/IWork/extract-date-formats.py` with opt-in numbers-parser 4.19.0 under the C time locale. The manifest records the extraction platform and a calendar-year display probe because the provider uses platform `strftime` behavior for early years. It verifies pinned source and native-export hashes, retains all 89 selected cells and their independent display values, and records localized native format codes and volatile cache differences separately. This evidence does not qualify native appearance, locale, calendars or time zones.

```sh
python Build/IWork/extract-date-formats.py OfficeIMO.TestAssets/Documents/IWorkCorpus OfficeIMO.TestAssets/Documents/IWorkCorpus/numbers-parser/date-formats.json
```

`numbers-parser/duration-ranges.json` retains 20 selected duration declarations and saved caches from the existing hash-pinned corpus: 18 day-only cells and two week/day/hour cells. Whole-day values, including negative values, have independent parser display evidence. Fractional-day rounding and native export/appearance equivalence remain unqualified; week/day/hour declarations remain unassessed by the destination adapters. Regenerate with the pinned opt-in provider:

```sh
python Build/IWork/extract-duration-ranges.py OfficeIMO.TestAssets/Documents/IWorkCorpus/numbers-parser OfficeIMO.TestAssets/Documents/IWorkCorpus/numbers-parser/duration-ranges.json
```
## Native Numbers fixed second ranges

[`native-exports/numbers-duration-seconds-v14.5.json`](native-exports/numbers-duration-seconds-v14.5.json) records a controlled Numbers 14.5 workbook, its Apple XLSX/PDF exports, three inspected PDF pages, and native Excel 16.113.3 observations on macOS 27.0.1. The OfficeIMO-authored intake is imported into Numbers; each C2:D8 range is explicitly set to fixed abbreviated hour/second, minute/second or second-only units, and D2:D8 contains native `C+C` formulas. Sources and exports are MIT-licensed OfficeIMO contributor fixtures. The native PDF embeds the host's Calibri font; OfficeIMO does not bundle it.

The generated decode retains all 42 selected declarations, raw seconds, 21 typed formula caches and matching native XLSX formats/day serials. Saved/reopened OfficeIMO XLSX matches those codes and caches; edited duration operands recalculate. Thirty integral display cases match the Numbers PDF. Fractional-second display remains approximated: Numbers rounds, the shared elapsed formatter can truncate, and native Excel differs by range, including hashes for the recorded negative hour/second and minute/second cells. The manifest records these limits instead of treating matching values or codes as proof of appearance. The default complete PDF workflow retains all populated rows and 21 formula results across six inspected pages with caller fonts and reported font/display approximation. Apple exports three positioned-sheet pages; the composed route splits the formula column onto each second page. Synthetic Pages/Keynote and blank-cell tests qualify adapter behavior; their native rendering remains separate.

Regenerate the decode with the pinned opt-in oracle; PDF tables, screenshots and native application observations are separate evidence:

```sh
python Build/IWork/extract-duration-seconds.py OfficeIMO.TestAssets/Documents/IWorkCorpus/native-exports/numbers-duration-seconds-v14.5.numbers OfficeIMO.TestAssets/Documents/IWorkCorpus/native-exports/numbers-duration-seconds-v14.5.xlsx OfficeIMO.TestAssets/Documents/IWorkCorpus/native-exports/numbers-duration-seconds-v14.5-decode.json
```

## Native Numbers fixed millisecond ranges

[`native-exports/numbers-duration-milliseconds-v14.5.json`](native-exports/numbers-duration-milliseconds-v14.5.json) records a task-owned copy of the preceding Numbers 14.5 reference with milliseconds enabled on all three C2:D8 ranges. The saved local source, Apple XLSX/PDF exports, three UI screenshots and three native PDF pages retain the same MIT contributor provenance. Independent decoding pairs all 42 raw-second values, 21 native formulas and five-field selected formats with Apple XLSX day serials and codes. The native PDF embeds host Calibri; OfficeIMO does not bundle it.

Numbers displays separate millisecond labels, such as `1h 1m 1s 750ms`. Its XLSX export uses fractional-second codes, including a terminal literal space; the shared formatter displays `1h 1m 1.750s `. The manifest records native PDF display and independently calculated shared display separately. Saved/reopened conversion preserves numeric caches and codes, edited operands recalculate, and the complete PDF route retains every populated row and all 21 formula results across six inspected pages. Native Excel appearance, millisecond-only totals, other label styles and automatic units remain unqualified. PDF text extraction misreads some label ligatures; visual inspection verifies the labels.

Regenerate the decode with the same isolated oracle:

```sh
python Build/IWork/extract-duration-seconds.py OfficeIMO.TestAssets/Documents/IWorkCorpus/native-exports/numbers-duration-milliseconds-v14.5.numbers OfficeIMO.TestAssets/Documents/IWorkCorpus/native-exports/numbers-duration-milliseconds-v14.5.xlsx OfficeIMO.TestAssets/Documents/IWorkCorpus/native-exports/numbers-duration-milliseconds-v14.5-decode.json
```

## Keynote slide background evidence

`keynote-backgrounds.json` records slide-style chains and background declarations from four unchanged pinned native Keynote packages and `native-exports/keynote-colors-v15.4.key`. It includes inactive templates; conversion tests compare only selected slides. Opaque RGB/sRGB colors qualify the additional fixed32 field `13` only at value `1`. Its meaning remains unspecified; other values and Display P3 remain unsupported. Existing fixture licenses and provenance above apply. Reproduce the manifest with numbers-parser 4.19.0:

```sh
python Build/IWork/extract-keynote-backgrounds.py OfficeIMO.TestAssets/Documents/IWorkCorpus OfficeIMO.TestAssets/Documents/IWorkCorpus/keynote-backgrounds.json
```

Saved/reopened PPTX verifies supported selected colors. Synthetic inputs cover explicit no-fill overrides, inheritance, malformed styles and fallback. `native-exports/keynote-colors-v15.4.json` separately records the native controlled sample, exports and rendering comparison. The sample derives from the licensed `nim-iwork` fixture, with a wrapped second-slide title and blue/red background presets authored and saved in Keynote 15.4. Both strict OfficeIMO slides match Apple's PPTX round-trip rasters at 72 dpi. Original Keynote-to-PPTX ligature differences remain, and broader fonts, themes, masters and producer coverage stay unqualified. Reproduce native exports with best-quality PDF, all slides and optional content/password off; render each recorded PDF with `pdftoppm -r 72 -png`.

## Native Keynote creation evidence

`native-exports/keynote-created-v15.4.json` pins an OfficeIMO C# generated package, its Keynote 15.4 PDF export and the native package saved after changing `café` to `Edited in Keynote`. These newly authored assets use the repository MIT license and contain no seed package or embedded font files. The executable creation example in `IWorkKeynoteWritingTests` generates static slides with two backgrounds, positioned Arial/Times New Roman text, Polish characters, emoji and a blank slide. The shared package smoke consumer exercises native creation and saved/reopened PPTX conversion through matching packed packages.

The independent pinned schema check requires initialized native records and an exact reference graph, validates metadata plists and deterministic ZIP fields, and compares modern text-fill colors with their legacy mirrors. Native open/export/edit/save/close/reopen qualifies this fixture on Keynote 15.4. PDF text, page sizes, embedded font names and 72 dpi RGB pixels provide separate rendered evidence. Reproduce the package through the documented creation API and export PDF with best image quality, all slides and optional content/password off. Modern native character and paragraph fills take precedence over legacy font colors; unsupported gradients, alpha, color spaces and no-fill declarations retain source diagnostics. Other fonts, scripts, application versions and source save-back remain outside this evidence.

Compile `KNArchives.proto` and `TSPArchiveMessages.proto` from the manifest's pinned schema revision with `protoc --include_imports --descriptor_set_out=<descriptor.pb> --proto_path=<proto-folder> KNArchives.proto TSPArchiveMessages.proto`. Run the opt-in isolated numbers-parser 4.19.0 check:

```sh
python Build/IWork/verify-keynote-creation.py OfficeIMO.TestAssets/Documents/IWorkCorpus/native-exports/keynote-created-v15.4.key <descriptor.pb> <evidence.json>
```

## Native Keynote rich-run creation evidence

`native-exports/keynote-created-rich-v15.4.json` pins the template-free three-slide creation model in `IWorkKeynoteRichTextWritingTests`, its Keynote 15.4 PPTX/PDF exports, all-slide rasters and the package saved after editing the plain run `plain` to `reset`. These newly authored assets use the repository MIT license. The two text boxes exercise Helvetica Neue, Arial and Times New Roman, 24/30/32-point sizes, opaque colors, bold/italic/single underline/single strikethrough, explicit plain resets, LF paragraphs, Polish text and a non-BMP emoji. The third slide stays blank.

Independent schema decoding checks all 41 created records, required fields, exact declared references, modern color mirrors and character boundaries in UTF-16. Plain rich ranges retain explicit false decoration values. The recorded native PPTX properties and saved/reopened OfficeIMO PPTX agree on complete text/run order, decoration, sizes and geometry; native font aliases and theme-backed Arial are recorded separately. Keynote opens without repair and preserves the native edit after save, close and reopen. Complete native PDF rasters were inspected; direct OfficeIMO PDF appearance remains a separate qualification route. No standalone fonts are distributed.

Reproduce the created package through `RichExample` in the executable writing test, then use the existing pinned-schema verification command above. Extract the selected Apple-saved character tables and native PPTX properties with the isolated numbers-parser 4.19.0 environment:

```sh
python Build/IWork/extract-keynote-rich-runs.py OfficeIMO.TestAssets/Documents/IWorkCorpus/native-exports/keynote-created-rich-edited-v15.4.key OfficeIMO.TestAssets/Documents/IWorkCorpus/native-exports/keynote-created-rich-v15.4.pptx <ranges.json>
```

## Native Keynote image creation evidence

`native-exports/keynote-created-images-v15.4.json` pins the `ImageExample` model in `IWorkKeynoteImageWritingTests`, its native Keynote 15.4 PPTX/PDF exports and all-slide rasters. Three 96 × 64-pixel inputs exercise opaque PNG, PNG alpha 0/128/255 and JPEG stretching. Text/image insertion order includes an obscured label and text in front of the first image. Native accessibility descriptions include a non-BMP emoji. The repository MIT license covers these newly authored documents and images.

The pinned independent decoder checks image-class extension 100, object/data references, payload checksums and reserved metadata IDs. Apple and saved OfficeIMO PPTX retain identical image payloads, frames, descriptions and stacking. The edited native package records a JPEG position change from 60 to 80 points, saved, closed and reopened in Keynote. Direct OfficeIMO PDF inspection covers all three pages, complete visible text, images and stacking with explicit default-policy font substitution; native text-layout equivalence remains unqualified. Empty trailing paragraphs do not report clipped visible content. The [macOS NativeAOT receipt](../../../Docs/Compatibility/iwork-keynote-creation-native-aot-macos-arm64-2026-10-09.json) records identical packages on .NET 8 and 10 using the fixed encoded inputs. Use `verify-keynote-creation.py` with the same pinned descriptor to reproduce graph qualification.

### Integer-function compatibility boundaries

`numbers-parser/integer-function-boundaries.json` records two unchanged formulas in
`single-cell-formulas.numbers`: `GCD(128,80,44,2^53)` and
`LCM(128,80,44,2^53)`. Independent integer arithmetic agrees with their retained
producer caches within relative tolerance 1e-14. The existing source checksum,
provenance and MIT notice apply.

[Excel GCD](https://support.microsoft.com/en-us/excel/functions/gcd-function)
documents `#NUM!` for an operand at or above 2^53;
[Excel LCM](https://support.microsoft.com/en-us/excel/functions/lcm-function)
documents that error for a result at or above 2^53. These are documented destination
contracts, not live Excel execution evidence. The source reader retains the numeric
caches without claiming complete editable reconstruction of these expressions.
Native recalculation, export and broader coercion behavior remain unqualified.

Reproduce with the pinned numbers-parser 4.19.0 environment:

```sh
python Build/IWork/extract-integer-function-boundaries.py OfficeIMO.TestAssets/Documents/IWorkCorpus/numbers-parser/single-cell-formulas.numbers OfficeIMO.TestAssets/Documents/IWorkCorpus/numbers-parser/integer-function-boundaries.json
```

The [controlled Numbers 14.5 reference](native-exports/numbers-integer-functions-v14.5.json)
adds 17 cases, 70 formulas, Numbers XLSX/PDF exports and a separate Excel 16.113.3
copy saved after **Calculate Now**. The authored intake, native source, both exports
and screenshots are hash-pinned under the repository MIT license. The independent
decoder records function identities 58 (`GCD`) and 75 (`LCM`), source values,
formula text and both producers' saved results.

Nineteen literal, twelve populated numeric-range and six numeric/blank-range calls qualify reconstruction and local recalculation. All-blank, mixed-type, formula-dependent and whole-axis references,
text, Booleans and large operands retain the narrower source reconstruction
boundary. Native Numbers stores a numeric-text range LCM as zero where Excel
recalculates 360. Native `2^53` and `2^53-1` power caches both round to
`9007199254740990`; fresh Excel caches preserve the distinct integer results.
The unchanged large-operand fixture above retains its separate evidence limits.
Excel 16.113.3 accepts referenced `2^53` in GCD despite the documented error
boundary. Neither that observation nor matching function names qualifies a wider
domain. Native Numbers errors still have no XLSX error code; their PDF cells are
blank. The complete native PDF and both app screenshots were inspected.

The [native reference-edit fixture](native-exports/numbers-integer-references-v14.5.json)
changes B2 from 24 to 30 and replaces two literal calls with `GCD(B2,C2,D2)` and
`LCM($B$2,$C$2:$D$2,1)`. Numbers 14.5 retains 6 and 180 after save/close/reopen.
The pinned independent decoder checks fourteen reference calls against the native
XLSX and integer arithmetic. Saved OfficeIMO XLSX retains the expressions and
recalculates to 12 and 360 after restoring B2 to 24. Native screenshots and the
one-page PDF are inspected. The separate [cross-table reference](native-exports/numbers-integer-cross-table-v14.5-decode.json) qualifies eight numeric calls across tables and sheets. Complete PDF reconstruction and broader coercion remain unqualified.

```sh
python Build/IWork/verify-numbers-integer-references.py OfficeIMO.TestAssets/Documents/IWorkCorpus/native-exports/numbers-integer-references-v14.5.numbers OfficeIMO.TestAssets/Documents/IWorkCorpus/native-exports/numbers-integer-references-v14.5.xlsx integer-references-decode.json
```

The [integer-text reference](native-exports/numbers-integer-text-reference-v14.5.json) records four direct scalar calls with a text operand, `"12"`, changed natively to text `"18"`. Numbers 14.5 retains the text type and dependent results after save/close/reopen. The pinned independent decoder verifies native single-cell AST nodes, table UUIDs and caches against native XLSX and integer arithmetic. Saved OfficeIMO XLSX preserves text, formulas and recalculation. The native one-cell range entries normalize to single-cell nodes; text rectangles and colon ranges retain their separate incomplete-expression boundary. Six blank-reference controls retain their formulas; four Boolean and arbitrary-text controls remain partial. This reference does not qualify complete PDF or appearance equivalence.

```sh
python Build/IWork/verify-numbers-integer-text.py OfficeIMO.TestAssets/Documents/IWorkCorpus/native-exports/numbers-integer-text-reference-v14.5.numbers OfficeIMO.TestAssets/Documents/IWorkCorpus/native-exports/numbers-integer-text-reference-v14.5.xlsx OfficeIMO.TestAssets/Documents/IWorkCorpus/native-exports/numbers-integer-text-reference-changed-v14.5.numbers OfficeIMO.TestAssets/Documents/IWorkCorpus/native-exports/numbers-integer-text-reference-changed-v14.5.xlsx integer-text-reference-decode.json
```

The [blank-reference pair](native-exports/numbers-blank-reference-v14.5-decode.json) contains ten numeric/text/blank calls and four all-blank direct error controls. It is a controlled Apple import of the repository-owned text-reference XLSX, with explicit blank-only controls, saved and reopened in Numbers 14.5. A native edit fills C4 with numeric 18; the second source and both native XLSX exports retain the resulting fourteen numeric calls. Independent raw AST and tile row indexes preserve the leading unstored row and bind coordinates through table UUIDs. The decoder's public presentation strings are not used as the coordinate oracle. Direct `LCM(blank,6)` yields 6; the numeric/blank range yields 0, then 90 after the edit. Saved OfficeIMO XLSX reproduces the ten baseline expressions and their edited values; the four omitted error formulas keep their source error caches. A separate native Excel 16.113.3 fill/clear observation agrees. Both native reopen screenshots were inspected. This reference does not qualify appearance, complete PDF, native error codes or other producer versions.

```sh
python Build/IWork/verify-numbers-blank-references.py OfficeIMO.TestAssets/Documents/IWorkCorpus/native-exports --out numbers-blank-reference-decode.json
```

Reproduce the controlled fixture decode with the pinned oracle:

```sh
python Build/IWork/extract-numbers-integer-functions.py OfficeIMO.TestAssets/Documents/IWorkCorpus/native-exports/numbers-integer-functions-v14.5.numbers OfficeIMO.TestAssets/Documents/IWorkCorpus/native-exports/numbers-integer-functions-v14.5.xlsx OfficeIMO.TestAssets/Documents/IWorkCorpus/native-exports/numbers-integer-functions-excel-v16.113.3.xlsx OfficeIMO.TestAssets/Documents/IWorkCorpus/native-exports/numbers-integer-functions-v14.5-decode.json
```

The [Excel optional-lookup reference](native-exports/excel-xlookup-optional-v16.113.3.json)
records seven XLOOKUP cases after explicit native recalculation from an intake
without formula caches. Empty fallback, match-mode and search-mode positions
retain their native defaults; reverse search and explicit fallback have separate
controls. The original intake, saved Excel 16.113.3 workbook and inspected app
screenshot are hash-pinned. This qualifies these optional-argument cases only.

### Complete native-content PDF references

The [paragraph-layout creation reference](native-exports/pages-created-layout-v15.4.json)
also retains its composed one-page PDF, inspected raster and public workflow
result. Independent extraction agrees with all 75 native words in order. The
eight selected source paragraphs are reconstructed with no omitted or unassessed
content units; opaque record fidelity remains unassessed. Spacing, custom tabs,
blank-paragraph height and rich/default resets remain visible. PDF font inspection
confirms 14-point Helvetica Bold for `Rich layout` and 12-point Helvetica for the
plain run. The shared portable profile reports Arial substitution and retains its
separate typography and monochrome-emoji limits.

`native-exports/iwork-composed-pdf-portable.json` pins the Apple-edited Pages body,
Numbers scalar workbook and Keynote table deck, all three composed PDFs, nine
composed page rasters and their native comparison exports. The fresh Keynote PDF
includes the edited numeric/text cells and text entered into the formerly empty
table. `keynote-created-tables-pdf-export-v15.4.key` separately retains the
metadata/resource normalization performed by Keynote during that export; its
composed pages match the original edited fixture. Original source fixtures remain
unchanged. These authored documents use
the repository MIT license; the registered and embedded fallback fonts retain
their recorded SIL OFL 1.1 licenses.

The workflow uses the public shared rendering profile with four Carlito style
faces, Noto Emoji and the Noto Sans JP common subset, under the portable
deterministic PDF resource policy. It discovers no installed fonts. The manifest
records font checksums, producer versions, selected-unit coverage and both
stages' fidelity diagnostics. Independent Poppler text extraction verifies all
80 numbered Pages paragraphs in order, every populated Numbers scalar/literal
value and all Keynote text and native edits. Every composed and comparison page
is visually inspected; the completely unset Numbers grid retains its drawn
page. The workflow tests exercise this contract on .NET 8 and 10.

The composed exports have two Pages pages, four Numbers pages and three Keynote
pages. They retain complete selected content with visible font, text-inset,
monochrome-emoji and invariant/localized display differences. Pages breaks its
body after paragraph 42 instead of Apple's 41. Numbers follows the XLSX
table-per-worksheet projection and splits the wide first table by columns;
Apple's positioned-sheet export has two pages. Opaque source-record fidelity,
native typography/pagination equivalence, additional scripts/producers and
composed-PDF NativeAOT/platform acceptance remain separate qualification work.
Use the [workflow font example](../../../OfficeIMO.Workflows.IWork/README.md) and
the manifest's routes and reproduction commands to regenerate and inspect the
outputs.

## Native Pages inline-image creation

[`native-exports/pages-created-images-v15.4.json`](native-exports/pages-created-images-v15.4.json) records OfficeIMO-authored PNG/JPEG body fixtures under the repository MIT license. Independent pinned schemas verify the mixed and image-only graphs, resources, pixel metadata and native-saved attachment/frame values. Pages 15.4 opens the created files without repair, resizes one image, saves and reopens it, and exports DOCX/PDF references. Native DOCX frames and descriptions agree with strict OfficeIMO DOCX. Independently decoded PDF images retain exact source pixels and alpha; all native and composed pages are inspected. The manifest records font, emoji, vertical-gap and native PDF extraction limits. An edited native PDF, other Apple versions and broader media/placement profiles remain unqualified.

## Native Pages decimal-list creation

[`native-exports/pages-created-numbering-v15.4.json`](native-exports/pages-created-numbering-v15.4.json) pins OfficeIMO-authored mixed and list-only documents under the repository MIT license. The mixed body contains two decimal levels, starts/restarts, parent resets, plain text, a scalar table, a transparent PNG and return to an earlier group. Independent schemas decode every created record and selected native-saved body/style declaration. Pages 15.4 opens both documents without repair; an inserted parent survives save, close and reopen. Native DOCX indents and all 14 original/15 edited PDF markers agree with strict OfficeIMO DOCX and composed PDFs. All native and composed pages are inspected.

Reproduce the created graphs with `Build/IWork/verify-pages-creation.py` and the native-saved declarations with `Build/IWork/extract-pages-list-profile.py`, using the pinned descriptor digest and commands in the manifest. The native edit demonstrates sparse restart metadata: a level persists into an inserted paragraph, while its restart flag applies only at the original boundary. Screenshots and PDF exports supply visible counter evidence because Pages accessibility text reports different ordinal values. The manifest records font, emoji, spacing, native PDF extraction and group-restoration limits; additional list profiles and receiving-application OfficeIMO DOCX appearance remain unqualified.

The [rich-item reference](native-exports/pages-created-rich-numbering-v15.4.json) adds bold, italic, underline, strikethrough, opaque colors and explicit plain resets at the same two decimal levels. The 51-record graph and native-saved character ranges independently retain all eight counters and plain black marker defaults. A final text edit survives Pages 15.4 save/close/reopen. Strict OfficeIMO DOCX and complete composed PDFs retain the rich runs; all native and composed PDF pages are inspected. Use the same independent graph/style tools with the manifest's descriptor. Native DOCX retains text/counters but changes font families and selected bold/italic overrides. Mixed fonts/sizes, broader lists, exact typography, emoji and pagination equivalence remain unqualified.

The [third-level reference](native-exports/pages-third-level-v15.4.json) qualifies plain Arial 12 items at levels 0, 1 and 2, ancestor resets, starts/restarts, plain gaps and return to an earlier group. Independent decoding and native DOCX/PDF exports retain all 16 original and 17 edited counters. A grandchild insertion survives Pages 15.4 save/close/reopen. The saved OfficeIMO DOCX has identical word bounds across all 74 words when rendered by Pages; complete composed PDFs retain content with explicit portable-font substitution. All six PDF pages are inspected. Reproduce with `Build/IWork/verify-pages-third-level.py <native-exports> --descriptor <descriptor.pb> --out <evidence.json>` and the pinned descriptor in the manifest. This reference covers three levels. Tiered numbering, wider typography, pagination and additional producers remain separate qualification gates.

The [decimal-item edit reference](native-exports/pages-preserved-decimal-v15.4.json) qualifies four fixed-length replacements across levels 0, 1 and 2 in the Apple-saved third-level source. The existing raw text-preservation verifier accounts for all 13 entries and 98 messages, preserving every byte outside the declared text ranges. Pages 15.4 opens the result without repair and retains a further list-item edit after save/close/reopen. Native DOCX/PDF and strict OfficeIMO DOCX retain all 16 counters and indents. The Pages-rendered OfficeIMO DOCX matches all 74 native word bounds in this Arial 12 reference; all five PDF pages are inspected. Complete composed PDFs retain text and counters with explicit font substitution. Reproduce with `Build/IWork/verify-pages-decimal-editing.py <native-exports> --descriptor <descriptor.pb> --out <evidence.json>` and the pinned descriptor in the manifest. Length/structural changes, mixed runs, additional list profiles, broader typography and cache regeneration remain unqualified.

The [nine-level reference](native-exports/pages-nine-level-v15.4.json) qualifies all native decimal levels 0–8, ancestor resets, starts/restarts and resumed scopes in one Arial 12 body document. Nine fixed-length replacements preserve every byte outside their declared text ranges across 13 package entries and 120 IWA messages. Pages 15.4 retains root, middle and deepest-item edits and a new deepest-level item after save/close/reopen. Native exports and strict saved OfficeIMO DOCX retain all 37 original and 38 native-edited counters; its Pages rendering matches all 155 native word bounds in the preserved reference. All nine native/composed/receiving PDF pages are inspected. Composed PDFs retain complete content with explicit font substitution; tiered numbering, mixed fonts/sizes, broader typography, structural edits and cache regeneration remain outside this reference. Reproduce with `Build/IWork/verify-pages-nine-levels.py <native-exports> --descriptor <pages-descriptor.pb> --out <evidence.json>` using the pinned Pages descriptor and isolated numbers-parser 4.19.0. This is an opt-in qualification tool with no runtime dependency.

## Native numeric-format creation reference

[`native-exports/iwork-created-numeric-formats.json`](native-exports/iwork-created-numeric-formats.json) pins newly authored MIT tables with explicit decimal, percentage and scientific formats. Each of Numbers 14.5, Pages 15.4 and Keynote 15.4 opens its independently decoded template-free graph, recognizes the format families, retains three numeric edits after save/close/reopen and exports native editable and PDF references. All 57 typed cells and formats survive; all three PDF pages are inspected. The manifest records source/export hashes, precision, grouping, negative styles, locale, table-placement differences and qualification limits. Native PDFs embed application fonts; no standalone font is redistributed.

Reproduce each created graph with its family-specific `Build/IWork/verify-*-creation.py` command and pinned descriptor. Decode the selected formats after native edits with `Build/IWork/extract-native-table-formats.py <package> --descriptor <descriptor.pb> --out <decode.json>` using isolated numbers-parser 4.19.0. These are opt-in qualification tools, with no production dependency. Executable creation and saved-destination checks live in `IWorkNumericFormatNativeExportTests`; the smoke runner uses the same manifest under actual NativeAOT execution.

## Native Numbers calculation registration

The [calculation reference](native-exports/numbers-created-calculation-v14.5.json) pins the current scalar creator and a two-sheet, three-table input with complete native table formula-owner registrations. Numbers 14.5 opens both without repair and calculates newly entered formulas. The [independent call checks](native-exports/numbers-integer-cross-table-v14.5-decode.json) compare eight numeric cross-table GCD/LCM calls and an arithmetic control against native XLSX and integer arithmetic. A native operand edit changes the dependent caches; all three native PDF pages are inspected. The manifest records fonts, default print pagination, source/export hashes and the existing XLSX padding and placement limits. Historical creation and format references retain their original packages.

Use `Build/IWork/verify-numbers-creation.py --require-calculation-owners` with the pinned Numbers descriptor for the current created graphs. Historical scalar-only references omit that flag and report their missing calculation registrations. Reproduce the call evidence with `Build/IWork/verify-numbers-integer-cross-table.py <source.numbers> <native.xlsx> <changed.numbers> <changed.xlsx> <output.json>` in the isolated numbers-parser 4.19.0 environment. These tools are opt-in evidence runners with no product runtime role.
