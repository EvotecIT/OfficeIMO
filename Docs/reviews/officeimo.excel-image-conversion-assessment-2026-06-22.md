# OfficeIMO Excel To Image Conversion Assessment

Date: 2026-06-22

Roadmap: `Docs/officeimo.image-conversion-roadmap.md`

Current branch status: the first Excel image-export vertical slice now exists. Ranges, worksheets, and workbooks can export to PNG/SVG through `OfficeIMO.Drawing`; the shared raster stack owns PNG read/write, alpha blending, supersampled storage/resolve, polygon and even-odd contour fills, line/polyline/styled strokes, dashed ellipse stroke approximation, solid elliptical arcs, rotated ellipse fill/stroke, rotated/scaled image drawing, anchored text-line rendering, shared plain text-block raster/SVG rendering, shared SVG text/tspan writing, fallback glyph rendering, cached text measurement, rectangular clipping scopes, adaptive coverage sampling for supersampled render targets, basic text rendering, point distance, polyline-by-length interpolation, raster path flattening for line/quadratic/cubic path commands, supported SVG path-command serialization, shared SVG linear-gradient definition emission, shared raster linear-gradient rectangle fills, and shared Office/Visio dash vocabulary normalization. Excel image export now has approved PNG/SVG visual baseline gates, styled cell font sizes, basic shrink-to-fit, basic numeric text rotation through shared Drawing, explicit vertical alignment with Excel-like bottom alignment when the source style is unset, SVG and PNG cell text clipping, non-rotated plain cell text rendering through the shared Drawing text-block renderer, single-line, hard-break, shrink-to-fit, and basic rotated rich text run rendering through shared rich text block layout, direct/theme/indexed color resolution with tint/shade support for cell fills/fonts/borders/sheet tabs, deterministic hatch approximations for Excel pattern fills, simple two-stop linear gradient cell fills through shared Drawing, styled Excel borders for solid, hairline, medium, thick, dashed, dotted, dash-dot, dash-dot-dot, double-line, and diagonal border cases, worksheet hyperlink visual hints with an opt-out flag, rendered Excel-style comment indicators with unsupported-body diagnostics, first-pass simple worksheet drawing-object rendering through shared Drawing, range-clipped worksheet images whose visual rectangle overlaps the selected range even when anchored just outside it, basic authored worksheet picture rotation, first-pass conditional color scales/data bars, bounded numeric cell-is and simple comparison formula differential fills with priority and stop-if-true behavior, stable warnings for unsupported conditional rule/formula shapes, stable pattern-fill approximation and complex/path/unresolved gradient-fill unsupported diagnostics, source-referenced comment/note and threaded-comment metadata through a resolver shared by inspection, feature reporting, and image export, source-referenced worksheet drawing-object classification through a resolver shared by PDF preflight and image export, a source-ordered drawing layer for supported mixed shapes/images/charts, explicit hidden row/column omission behavior with `IncludeHidden`, worksheet print-area export through `UsePrintArea`, workbook print-area orchestration through `UseWorksheetPrintAreas`, image byte-format detection with SVG embedding for known compatible formats, and a shared Excel image/autofit number-format display helper for percent/date/number display cases plus custom literal affixes, escaped literal characters, and positive/negative/zero format sections. The Excel baselines now cover a polished range with image/chart/rich-text/comment-indicator content plus dedicated rich-text, conditional-formatting, sparkline, drawing-object, clipped-image, two-cell image, cropped-image, and rotated-image ranges; focused contract tests cover default bottom vertical alignment, custom number-format literal/section display text, shared Drawing plain text-block raster/SVG rendering, shared Drawing SVG text/tspan writer output, rotated PNG text clipping, hyperlink hint SVG output, range hyperlink expansion in inspection snapshots, single-line rich text SVG styling, hard-break rich text PNG/SVG preservation, shrink-to-fit rich text PNG/SVG preservation, basic rotated rich text run preservation with rotation diagnostics, deterministic Drawing text measurement, Drawing raster clipping, Drawing linear-gradient raster fills, Drawing polyline interpolation, shared Visio line-pattern and Office preset-dash normalization, OpenXML pattern fill SVG/PNG rendering with diagnostics, simple OpenXML gradient fill SVG/PNG rendering plus unsupported-gradient diagnostics for unresolved cases, source-filtered rendered comments/threaded-comment indicators with unsupported-body diagnostics, supported rounded-rectangle drawing-object output through shared Drawing, mixed shape/image z-order in both directions with decoded PNG pixels and SVG order assertions, range-clipped overlapping worksheet image output, two-cell marker sizing, authored image crop, authored image rotation, and source-filtered worksheet drawing-shape diagnostics. Excel, Visio, and PDF raster visual-baseline comparisons now share one Drawing-backed PNG decode/encode/diff helper. `OfficeIMO.Visio` still has important private rendering adapter code, but the duplicate pixel, text-block rendering, SVG text/tspan writing, path math, and dash mapping engine is being actively retired into shared Drawing; PNG/SVG connector labels and collision-aware label layout now use shared Drawing geometry, Visio native PNG text now paints through the shared Drawing text-block renderer, Visio SVG text now writes through the shared Drawing text/tspan writer, and supported built-in Visio SVG artwork/database paths now use shared Drawing path-command serialization instead of local `M`/`L`/`Q`/`C`/`Z` string assembly. Native Visio premium PNG baselines were refreshed after reviewing the shared renderer output, and the native premium baseline gate now runs as per-scenario theories against those refreshed artifacts.

Latest architecture guard checkpoint: `OfficeIMO.Tests.DrawingArchitectureTests` now pins the shared-brain contract in executable form. The guard proves `OfficeIMO.Drawing` has no package or project references, primary image/PDF export owners reference `OfficeIMO.Drawing`, production image-rendering owner code does not use `System.Drawing`, SixLabors/ImageSharp, SkiaSharp, or ImageMagick namespaces, and Visio PNG remains a thin adapter over `OfficeRasterRenderTarget`, `OfficeRasterCanvas`, shared text-block rendering, and `OfficePngWriter`. This does not make Excel export premium by itself, but it prevents future work from quietly rebuilding a second image-rendering stack.

Latest Excel shared render-plan checkpoint: `OfficeIMO.Drawing.OfficeTextBlockRenderPlan` now creates measured plain and stacked text plans for left/top rectangles as well as center-based Visio text boxes. Excel PNG and SVG plain cell text both call one Excel policy adapter that maps cell style, rotation, wrapping, vertical alignment, and shrink-to-fit settings into that shared Drawing plan before renderer-specific output. The remaining Excel-owned parts are explicit: workbook style interpretation, rich-text policy, rotation approximation diagnostics, clipping diagnostics, and source references stay in the Excel adapter, while reusable layout, fit, and placement stay in Drawing.

Latest rich-text SVG checkpoint: Drawing now has `OfficeTextBlockRenderer.AppendSvgRichTextBlock` as the SVG companion to `DrawRasterRichTextBlock`. Excel cell rich-text SVG and header/footer formatted-run SVG now route through the shared Drawing helper for line top, baseline, cursor, style, decoration, optional rotation emission, and preservation of authored boundary whitespace. `OfficeTextBlockRenderer` emits `xml:space="preserve"` for SVG text segments that contain leading, trailing, or repeated whitespace so run-boundary spaces survive in browser/viewer rendering instead of making rich text look hand-spliced. Focused Drawing and Excel rich-text contracts prove the shared behavior, and the affected rich-text, text-spill, and premium-range SVG baselines were refreshed through the Excel image visual gate.

Latest data-bar geometry checkpoint: Drawing now owns the reusable resolved geometry for proportional data bars. Excel conditional-formatting image export continues to use `OfficeDataBarRenderer` for PNG/SVG output, while PDF table data bars and Visio shape-data graphics now consume the same ratio/clamping geometry before emitting native PDF rectangles or VSDX shapes. This removes another private bar-placement brain while preserving adapter-owned source semantics and native output contracts.

Latest Bezier geometry checkpoint: Drawing now owns reusable quadratic and cubic Bezier point sampling through `OfficeGeometry`. The shared Drawing path flattener and Visio preserved-geometry renderer both call that primitive, so Drawing raster output and Visio PNG/SVG export no longer maintain parallel Bezier formulas. Visio still owns VSDX geometry row parsing and format-specific curve types; Drawing owns the generic sampled-curve math future Office shape renderers can reuse.

Latest DrawingML preset-geometry consolidation checkpoint: `OfficeIMO.Drawing.OfficeShapePresets` now owns richer preset geometry for heart, cloud, donut, can, cube, and left-right arrow. Word native PDF rendering consumes the shared preset table through serialized OpenXML `prst` tokens instead of carrying a second private shape-geometry switch, while Word/PDF keeps WordprocessingML extraction, dimensions, and style application. This gives PowerPoint PDF, Word PDF, and future Excel object rendering one central preset vocabulary; focused Drawing and Word/PDF tests prove the shared geometry and rendering behavior on `net8.0` and `net472`.

Latest Excel worksheet DrawingML preset checkpoint: Excel worksheet drawing-object image export now validates authored `a:prstGeom` tokens through `OfficeIMO.Drawing.OfficeShapePresets`, carries the serialized preset name plus horizontal/vertical flip flags and rotation in the neutral visual snapshot, and renders supported preset shapes through the shared Drawing geometry table. Excel no longer has a separate rectangle/rounded-rectangle geometry switch for supported worksheet shapes; it still owns worksheet anchors, DrawingML extraction, fill/outline/text policy, and source diagnostics for unsupported colors, groups, connectors, and missing or unsupported geometry. Rotated supported preset shapes now attach an `OfficeTransform` to the shared `OfficeShape`, render through `OfficeDrawingRasterRenderer` and `OfficeDrawingSvgExporter`, and expand the worksheet overlay scene so rotated geometry is not clipped by the unrotated shape bounds. Rotated shape text now passes authored rotation into shared Drawing scene text and emits `ExcelDrawingShapeTextRotationApproximation` because text-box metrics remain non-Excel-exact. Focused object tests prove a shared `heart` path shape renders to PNG/SVG without `ExcelDrawingShapeUnsupported`, preserves flip metadata, emits an SVG path, carries authored rotation, paints decoded PNG fill pixels through the public range export path, and reports rotated text approximation diagnostics when text is present.

Latest rotated preset visual checkpoint: the shared `OfficeShapePresets` heart geometry was refined and a dedicated approved Excel visual baseline now gates a rotated DrawingML preset shape through the public range PNG/SVG export path. The baseline asserts SVG path, matrix transform, fill, and outline artifacts, rejects unsupported-shape and rotated-text diagnostics for the empty-text shape case, decodes the approved PNG for visible fill pixels, and was manually reviewed after fixture layout correction so the object does not collide with title or caption text. This is the current proof that Excel, Word/PDF, PowerPoint/PDF, and future object export all share one preset geometry/transform brain in `OfficeIMO.Drawing`; remaining premium object gaps stay explicit rather than hidden.

Latest rotated scene-text checkpoint: `OfficeDrawingText` now carries rotation metadata and an explicit rotation center, Drawing SVG scene text writes the shared text/tspan output with a rotate transform, rotated Drawing raster scene text paints through `OfficeTextBlockRenderer`, and Drawing quality analysis expands rotated text bounds. Excel worksheet drawing-object text now uses that central Drawing scene text path for authored shape rotation while retaining `ExcelDrawingShapeTextRotationApproximation` until Excel-exact text-box metrics, wrapping, and baseline parity are implemented. A dedicated approved rotated shape-text PNG/SVG baseline was manually reviewed and gates the public range export result.

Latest aligned scene-text checkpoint: `OfficeDrawingText` now carries vertical alignment metadata, and non-default scene-text placement routes through shared Drawing text-block rendering in both raster and SVG output. Excel worksheet drawing-object extraction now maps DrawingML paragraph alignment and `bodyPr` anchoring into the neutral visual snapshot, then passes those values to the shared Drawing text path instead of inventing an Excel-local placement loop. Focused Drawing tests pin right/bottom scene-text output, focused Excel object tests prove authored DrawingML alignment reaches public PNG/SVG export, and a dedicated approved aligned shape-text baseline was manually reviewed.

Latest shape-text bodyPr checkpoint: Excel worksheet drawing-object extraction now carries DrawingML `bodyPr` wrapping and text insets, including Office default inset values, into the neutral visual snapshot. The renderer uses those inset values to define the shared `OfficeDrawingText` rectangle instead of applying a private fixed padding heuristic, so shape text placement keeps moving toward Office-authored geometry while Drawing continues to own text emission. Focused Excel object tests prove default insets, authored zero insets, wrapped SVG line splitting, and PNG visible text on `net8.0` and `net472`.

Latest scene-text shrink checkpoint: `OfficeDrawingText` now carries opt-in shrink-to-fit metadata, and Drawing PNG/SVG scene text uses the existing shared single-line and wrapped fit-to-bounds layout primitives instead of clipping immediately at the authored font size. Excel worksheet shape extraction maps DrawingML `a:normAutofit` into that shared flag through the neutral snapshot, while Excel still owns the OpenXML bodyPr semantics. Focused Drawing and Excel object tests prove SVG font-size shrink plus PNG visible text on `net8.0` and `net472`.

Latest shape-text autofit diagnostic checkpoint: Excel worksheet drawing-object extraction now carries DrawingML `a:spAutoFit` into the neutral visual snapshot as an explicit resize-shape-to-fit-text flag. PNG and SVG export keep using the shared fixed-bound `OfficeDrawingText` scene-text path, but now emit `ExcelDrawingShapeTextAutoFitUnsupported` with a sheet/cell source reference instead of silently ignoring the authored shape-resize behavior.

Latest shape-text vertical-orientation checkpoint: Excel worksheet drawing-object extraction maps DrawingML `bodyPr vert` into a neutral `ExcelDrawingTextOrientation` value instead of leaking OpenXML enum spellings through the image snapshot. Simple `vert` shape text now routes through the shared `OfficeDrawingText` stacked-text path for PNG/SVG and is covered by a dedicated approved vertical shape-text baseline that was visually reviewed as a readable stacked label. Complex vertical variants such as `vert270`, East Asian, Mongolian, and WordArt orientations still emit `ExcelDrawingShapeTextVerticalOrientationUnsupported` with a sheet/cell source reference until their semantics are implemented deliberately.

Latest page-layout visual checkpoint: page-sliced worksheet PNG/SVG export now has a dedicated approved baseline for a physical Letter landscape page with repeated print-title rows, fit-to-width page setup, supported header/footer text chrome, stable formatting-approximation diagnostics, and decoded nonblank page content. This proves the public worksheet export path can compose page semantics onto the shared Drawing canvas, while automatic multi-page fit pagination, large-sheet tiling, broader paper-size coverage, and Excel-exact header/footer images remain explicit premium gaps.

Latest header/footer image checkpoint: page-sliced worksheet export renders `&G` header/footer images through shared Drawing composition instead of rejecting picture placeholders as unsupported chrome. Multiple header/footer image positions are preserved in the VML header/footer drawing part, SVG output embeds direct-safe images or transcodes through the shared/caller codec path, and raster output uses the same shared decoder policy. Undecodable sources produce a visible placeholder and `IMAGE_SOURCE_DECODE_FALLBACK` rather than an Excel-only skip diagnostic. A dedicated approved PNG/SVG header-footer-image visual baseline gates the public worksheet export path; placement remains a diagnosed `ExcelHeaderFooterImageApproximation` until Excel-exact image geometry is implemented.

Latest MIME consolidation checkpoint: `OfficeIMO.Drawing.OfficeImageInfo` now owns MIME parameter stripping and known image content-type alias canonicalization through shared helpers. Excel header/footer images, template image values, and URL image downloads consume the shared Drawing normalization instead of carrying local MIME cleanup rules, while header/footer and downloader paths still reject non-image responses at the adapter boundary. `OfficeSvgImageRenderer` also routes its internal MIME cleanup through `OfficeImageInfo`, so SVG image projection and Excel image package inputs no longer maintain separate normalization brains. Focused Drawing, Excel header/footer, downloader, and template tests prove `image/jpg; charset=binary` canonicalizes to `image/jpeg`, unknown `image/*` values remain accepted where the adapter contract allows them, and non-image MIME values are rejected by the shared helper.

Latest radar chart consolidation checkpoint: shared `OfficeChartDrawingRenderer` now emits complete radar series as closed polygon shapes with shared fill, stroke width, and stroke dash metadata instead of splitting unfilled radar outlines into separate line segments. `OfficeRasterCanvas` now exposes a styled polygon stroke primitive and `OfficeDrawingRasterRenderer` routes polygon strokes through it, so authored dash styles are shared by PNG and SVG drawing output. Focused chart tests prove unfilled radar series remain unfilled, retain the authored stroke dash style, and emit SVG dash attributes; the full `PdfDocumentChartDrawingTests` class passes again on net8.

Latest checkpoint: authored worksheet sparklines are now found through a shared Excel sparkline resolver used by feature reporting, PDF preflight, and image export. Image export now renders visible same-sheet numeric line, column, and win/loss sparklines in PNG/SVG through the neutral snapshot and shared Drawing primitives, including basic authored colors, markers, negative coloring, and zero-axis output. Rendered sparklines emit `ExcelSparklineRenderingApproximation`; cross-sheet and unresolved sparkline data still emit precise source diagnostics. A dedicated approved PNG/SVG sparkline baseline now gates line, column, and win/loss output.

Latest object checkpoint: comments/notes and threaded comments now render visible cell indicators when their targets are inside the exported range. With `ShowCommentBodies` disabled, bodies/popovers keep `ExcelCellCommentUnsupported` / `ExcelThreadedCommentUnsupported` source diagnostics; with it enabled, visible classic and threaded bodies render as dependency-free approximation callouts with anchored pointers, enter the ordered drawing-layer overlay stream, paint callout body text through shared Drawing text-block emission, and emit `ExcelCellCommentBodyApproximation` / `ExcelThreadedCommentBodyApproximation`. Simple worksheet rectangle/rounded-rectangle drawing shapes plus supported DrawingML preset shapes with solid RGB fill/outline now flow through the neutral Excel visual snapshot and shared `OfficeIMO.Drawing` PNG/SVG renderers, including authored rotation for the shape geometry. Supported shapes, images, opt-in comment bodies, and charts now share one ordered drawing-layer overlay stream derived from worksheet drawing part order, so mixed object paint order is no longer split across separate renderer passes. Unsupported geometry, theme/system/transformed colors, connectors, groups, Excel-exact rotated text-box metrics, Excel-exact comment popover geometry/state, and non-chart graphic frames remain source-diagnosed, and a dedicated approved drawing-object baseline gates the first shape/text-box rendering slice. Focused object tests now prove both image-under-shape and shape-under-image cases with snapshot order, SVG order, and decoded PNG pixels, plus comment-body layer placement, SVG pointer/text/color output, decoded PNG callout pixels, rotated preset shape PNG/SVG output, and rotated shape text approximation diagnostics.

Latest image checkpoint: range image export now includes embedded worksheet images whose visual rectangle intersects the selected range even if the anchor cell is outside that range. PNG output clips through shared Drawing canvas bounds, hidden-anchor omissions remain diagnosed, two-cell anchored worksheet pictures now size from their OpenXML from/to marker geometry, and authored picture crop rectangles, basic rotation, and horizontal/vertical flips render in PNG/SVG from OpenXML transform metadata. SVG image projection now also uses shared `OfficeSvgImageRenderer` for source crop, clip paths, data URI emission, rotation, and flips instead of keeping that generic projection math in Excel. Excel still owns OpenXML extraction, content-type allow-listing, and source diagnostics. Dedicated approved clipped-image, two-cell image, cropped-image, rotated-image, and transformed-image baselines cover the behavior, and the cropped/two-cell/transformed artifacts were manually reviewed after this consolidation slice.

Latest chart checkpoint: Excel chart export now carries the first authored layout/style slice into the shared `OfficeIMO.Drawing` chart snapshot instead of rendering every worksheet chart with default chart styling. Chart/plot area solid fill, outline color, outline width, and preset outline dash, legend visibility/position/overlay, title overlay, axis titles, chart-level data-label flags/position/number format, simple authored series fill/line colors, line widths, and preset dashes, simple point fills, simple marker fill colors, marker visibility, simple marker size, simple marker solid outline color/width, simple circle/square/diamond/triangle/dash/dot/plus/X/star marker shapes, simple category/value major-gridline color/visibility/width/preset dash, simple category/value minor-gridline color/visibility/width/preset dash, simple category/value axis-line color/visibility/width/preset dash, simple category/value major tick marks, simple category/value minor tick marks, category/value axis label visibility when Excel tick-label position is `none`, simple high/low/next-to category/value tick-label placement, simple maximum-crossing horizontal category-axis and vertical value-axis placement, simple category/date-axis reverse-order rendering, simple value-axis number formats for vertical and horizontal bar orientations, simple value-axis display-unit scaling and labels, simple linear value-axis min/max/major/minor-unit scaling, simple title text color/font-family/font-size/bold/italic, simple legend text color, simple data-label text color, simple axis-label text color, simple axis-title text-color override, simple legend/data-label/axis-label font sizes, simple axis-title font size, simple legend/data-label/axis-label font families, simple legend/data-label/axis-label bold/italic buckets, and simple axis-title font-family/bold/italic overrides now map to shared `OfficeChartStyle`, `OfficeChartLayout`, and `OfficeChartSeries`. Trendlines, leader lines, point-level data-label overrides, picture markers and richer marker outline effects, custom/richer series dash/effect styling, richer point-level overrides, remaining Excel-exact minor-gridline/tick placement edge cases, custom dash/effect parity beyond preset gridline and axis lines, remaining tick-label placement edge cases beyond simple high/low/next-to/none, axis-crossing geometry beyond simple horizontal category-axis and vertical value-axis maximum crossing, log/value-axis-reverse-order/non-value-axis-unit/non-default cross-between axis geometry, Excel-exact display-unit placement/typography, full custom/date/scientific/conditional tick formatting including category/date-axis tick formats, richer chart title typography/effects beyond simple font-family/font-size/bold/italic, per-element chart text runs beyond the supported shared buckets, richer chart/plot area effects beyond simple solid fill/outline/width/preset dash, and richer series/data-label shape styling are still not premium-rendered, but they now emit stable diagnostics instead of disappearing. Focused SVG plus decoded PNG tests prove the supported simple chart/plot area border-width/dash, series-color, series-line-width, series-line-dash, point-color, marker-fill/size/shape/outline including line-based dash/dot/X/star markers, category/value major-gridline color/width/dash, simple value-axis minor-gridline color/width, category/value axis-line-color/width/dash, simple major and minor axis tick-mark rendering, title-color/title-typography, legend/data-label/axis-label and axis-title text colors, text-font-size including axis titles, text-font-family/style buckets including separate axis-title overrides, suppressed axis-label, supported value-axis-number-format paths, simple display-unit labels, simple linear value-axis min/max/major/minor-unit SVG output, and source-referenced diagnostics for remaining unsupported axis tick-label placement, remaining approximated minor tick-mark placement, remaining custom axis crossing, unsupported log/value-axis-reverse-order/non-value-axis-unit/cross-between axis scale, and category/date-axis number formats, and the premium approved range baseline was visually reviewed and refreshed for the improved chart output.

Latest chart-axis label checkpoint: shared `OfficeChartDrawingRenderer` now renders value-axis labels from the same major tick set used for major gridlines instead of defaulting unlabeled axes to only minimum and maximum labels. Excel image export receives this through the existing shared chart snapshot path, so worksheet chart SVG/PNG output now shows intermediate labels such as 500, 1,000, and 1,500 for the same chart that previously only showed endpoints. A dedicated approved chart-axis PNG/SVG baseline now gates this behavior, and the premium range baseline was refreshed after visual review.

Latest Excel text-overflow checkpoint: shared `OfficeTextLayoutEngine` now exposes an explicit overflow behavior so callers can choose ellipsis or caller-owned clipping without forking text layout. Excel cell text and rich text image export now request clip overflow through that shared path, so overflowing cell content remains the authored text and is clipped by the cell's PNG/SVG clip surface instead of being rewritten with a synthetic `...`. Focused plain/rich Excel tests prove clipping diagnostics and SVG output without ellipsis, while the premium and rich-text approved baselines were visually reviewed and refreshed. Premium still needs Excel-exact spill into adjacent empty cells and deeper baseline/font parity, but the current output is less hand-authored and closer to native worksheet clipping.

Latest consolidation checkpoint: shared text layout is no longer only a Visio PNG cleanup target. `OfficeIMO.Drawing` now owns `OfficeTextLayoutEngine`, `OfficeTextLine`, `OfficeTextBlockLayout`, `OfficeTextVerticalAlignment`, `OfficeTextPlacement`, `OfficeTextBlockRenderer`, `OfficeRichTextRun`, `OfficeRichTextSegment`, `OfficeRichTextLine`, and `OfficeRichTextBlockLayout` for dependency-free word wrapping, rich-run tokenization, long-word breaking, hard-break normalization, max-line measurement, single-line trim-to-width, shrink-to-fit font sizing for measured single-line and rich-run text, bounded wrapped text-block fit, bounded plain/rich text-block layout orchestration, visible-height text-block clipping with ellipsis, clipped-state reporting, reusable text top/anchor/line-left placement, shared point rotation for rotated text placement, measured plain text-block PNG/SVG emission with alignment, vertical placement, underline, and rotation support, and shared XmlWriter text/tspan output for SVG adapters. `OfficeIMO.Visio` native PNG and SVG text rendering both consume shared layout helpers, native PNG text consumes the shared Drawing text-block renderer, and Visio SVG text/connector labels now consume the shared Drawing text/tspan writer while keeping Visio-specific enum mapping, styling, label backgrounds, label-adjusted data markers, and coordinate mapping in the Visio adapter. Excel plain cell text rendering now consumes the shared line model and `LayoutTextBlock` coordinator for wrapped, hard-break, shrink-to-fit, clipped, and forced-single-line output in PNG/SVG, and non-rotated plain cell text now uses the shared Drawing text-block renderer for PNG/SVG emission; Excel rich cell text now consumes `LayoutRichTextBlock` for hard-break, wrapped, shrink-to-fit, and basic rotated run-preserving PNG/SVG output while still reporting rotation approximation diagnostics. This does not make Excel text premium by itself, but it removes another private renderer brain and gives wrapped/rich/rotated text hardening a shared foundation instead of parallel implementations.

Latest stacked-text checkpoint: `OfficeIMO.Drawing` now owns `OfficeTextLayoutEngine.LayoutStackedTextBlock` and `LayoutStackedRichTextBlock` for upright one-text-element-per-line stacked layout. Excel `TextRotation=255` now renders plain and rich cell text through that shared Drawing path in PNG/SVG and reports `ExcelCellTextRotationApproximation` instead of the old unsupported stacked-text or rich-layout fallback diagnostics. A dedicated approved stacked-text PNG/SVG baseline gates readable upright output, rich-run styling, nonblank PNG pixels, and SVG output without rotation transforms; premium still needs Excel-exact stacked baseline metrics.

Latest page-setup checkpoint: page-sliced worksheet image export now applies orientation, margins, manual scale, supported worksheet paper-size codes, and bounded one-page fit-to-width/fit-to-height scaling by composing the rendered worksheet page onto a physical page canvas through shared `OfficeImageComposer` PNG/SVG layers and shared Drawing physical page-size primitives. Excel keeps OpenXML page setup semantics and source diagnostics, while Drawing owns neutral page-size math and image composition. Missing paper size emits `ExcelPageSetupPaperSizeDefaulted`; unknown paper-size codes emit `ExcelPageSetupPaperSizeUnsupported` and fall back to Letter; fit-to-width/fit-to-height values above one page in either dimension still emit `ExcelPageSetupUnsupported` until real automatic multi-page fit pagination exists.

Latest Excel/PDF page-geometry consolidation checkpoint: `OfficeIMO.Drawing.OfficePageSize` now exposes point conversion alongside pixel conversion, and `OfficeIMO.Excel` owns `ExcelPageSetupGeometry` for worksheet paper-size resolution, fit-scale policy, and margin clamping. Excel image export and first-party Excel PDF now consume the same supported worksheet paper-size resolver instead of carrying separate maps; Excel PDF still lets explicit `ExcelPdfSaveOptions.PageSize` win, but otherwise honors supported worksheet paper-size codes such as A4. This removes another page-layout brain without moving PDF stream/page writer responsibilities out of `OfficeIMO.Pdf`.

Latest PDF image-placement consolidation checkpoint: `OfficeIMO.Drawing.OfficeImageRenderPlan` now owns reusable image target-box, fit, and source-crop placement math, including separate top-left and bottom-left coordinate-system entrypoints. The first-party PDF writer now consumes that Drawing plan for image draw rectangles in flow, table, header/footer, canvas, and page-background contexts instead of maintaining local crop-plus-fit or page-fit algorithms. PDF remains responsible for PDF clip paths, XObject streams, tagging, annotations, page ordering, opacity graphics states, and compression, but generic visual placement math has moved into the shared engine where Excel worksheet images and later adapters can reuse it. PDF link-annotation bounds for cover-fitted, clipped, or source-cropped images are now centralized in one PDF-owned helper shared by flow and table-cell image rendering.

Latest font diagnostics checkpoint: Excel cell, rich text run, chart text, and page-sliced header/footer image export now report explicit font-family fallback through `ExcelCellFontFamilyFallback`, `ExcelChartFontFamilyFallback`, and `ExcelHeaderFooterFontFamilyFallback` when a workbook requests a family that `OfficeIMO.Drawing` cannot load exactly. The probes use `OfficeTrueTypeFont.TryLoadFontFamily`, so font resolution remains centralized in Drawing while Excel keeps source references and adapter-owned diagnostic wording. Chart diagnostics ignore OpenXML theme placeholders such as `+mn-lt` because those are not literal font requests. Focused tests cover plain cell text, rich text runs, chart title text, and formatted header/footer text without regenerating visual baselines because the rendered pixels are intentionally unchanged.

Latest styles checkpoint: Excel pattern-fill rendering no longer owns the generic hatch drawing loops. `OfficeIMO.Drawing` now owns neutral hatch primitives through `OfficeHatchPatternKind`, `OfficeRasterCanvas.DrawHatchPatternRectangle`, and `OfficeSvgFormatting.AppendHatchPatternRectangle`; Excel keeps OpenXML pattern-name normalization, density policy, foreground/background resolution, clipping, and `FillPatternApproximation` diagnostics. A dedicated approved pattern-fill PNG/SVG visual baseline now covers horizontal, vertical, grid, diagonal, trellis, and dotted fallback output and was manually reviewed. This improves the "one rendering brain" goal, but pattern fills remain deterministic approximations rather than Excel-exact density/parity.

Latest sparkline consolidation checkpoint: Excel sparkline image export now uses shared Drawing primitives instead of maintaining separate PNG and SVG mini-chart geometry loops. `OfficeIMO.Drawing` owns `OfficeSparklineKind`, `OfficeSparklineStyle`, `OfficeSparklinePointStyle`, and `OfficeSparklineRenderer` for dependency-free line, column, and win/loss sparkline geometry/emission; Excel keeps OpenXML sparkline discovery, kind mapping, per-point color and marker policy, clipping, `ExcelSparklineRenderingApproximation`, and source diagnostics. The approved sparkline PNG/SVG baseline still passes without regeneration after the migration. Premium still needs Excel-exact hidden/empty data behavior, date-axis/group scaling, and broader sparkline variants.

Latest conditional-formatting consolidation checkpoint: Excel data-bar image output now uses `OfficeDataBarRenderer` in shared Drawing instead of local PNG/SVG rectangle math. Excel still owns conditional-formatting rule evaluation, start/width ratios, color resolution, unsupported icon-set diagnostics, source references, and all non-data-bar conditional-formatting semantics. The approved conditional-formatting PNG/SVG baseline still passes without regeneration, so the move removes another duplicate rendering loop without changing the reviewed output.

Latest Visio image consolidation checkpoint: `OfficeSvgImageRenderer` now has both `StringBuilder` and `XmlWriter` image-emission paths. Excel range SVG image output uses the builder path for worksheet picture crop/clip/rotation/flip projection, and Visio package-preview SVG artwork uses the writer path for data URI emission, preserve-aspect behavior, and rotation while keeping package preview discovery, metadata sniffing, and placement policy in Visio. Focused Drawing writer tests and existing Visio package-preview SVG tests prove PNG preview projection, generic metadata sniffing, content-type normalization, unsafe SVG fallback, and rotated preview artwork still work.

Latest Visio SVG primitive consolidation checkpoint: `OfficeSvgPrimitiveWriter` now owns simple XML SVG circle, rectangle, line, and path emission plus shared number/color/stroke-cap/stroke-join attributes. Visio built-in stencil metadata artwork keeps only stencil semantics, opacity, placement, and page/shape coordinate policy, and routes generic primitive output through the shared Drawing writer. Focused Drawing primitive-writer tests and existing Visio stencil SVG contracts prove vector metadata artwork and rotated stencil artwork still work.

Latest nested SVG consolidation checkpoint: `OfficeSvgFormatting` now owns inner SVG extraction and nested SVG viewport emission for child drawing fragments embedded into larger SVG exports. Excel chart overlays, supported drawing-object overlays, and opt-in comment body callouts now reuse that shared helper instead of each partial hand-building child `<svg>` wrappers and `viewBox` strings. Focused Drawing formatter tests plus Excel chart, drawing-object, comment-body, drawing-object baseline, and premium range baseline tests passed without baseline regeneration.

Latest SVG polygon consolidation checkpoint: `OfficeSvgFormatting` now owns complete SVG polygon element emission for `StringBuilder` renderers. Shared Drawing polygon shapes and Excel comment indicator/body pointer SVG output use the same point-list, fill, stroke, and stroke-width emission helper while preserving adapter-owned geometry and policy. Focused Drawing formatter/exporter tests plus Excel comment indicator, comment-body, threaded-comment, and premium range baseline tests passed without baseline regeneration.

Latest SVG line consolidation checkpoint: `OfficeSvgFormatting` now owns complete SVG line element emission for `StringBuilder` renderers. Shared Drawing line shapes and Excel range border SVG output use the same coordinate, stroke paint, opacity, width, dash-array, and line-cap helper while preserving adapter-owned transform and border-style policy. Focused Drawing formatter/exporter tests plus Excel border-style and premium range baseline tests passed without baseline regeneration.

Latest SVG rectangle consolidation checkpoint: `OfficeSvgFormatting` now owns complete SVG rectangle and rounded-rectangle element emission for `StringBuilder` renderers. Shared Drawing rectangle shapes, Excel range gridline and cell-fill SVG rectangles, shared data-bar SVG rectangles, and shared sparkline column/win-loss SVG rectangles use the same coordinate/size/corner-radius helper while preserving adapter-owned paint, transform, cell style, and conditional-formatting policy. Focused Drawing formatter/exporter/data-bar/sparkline tests plus Excel border-style, pattern-fill, conditional-formatting, sparkline, and premium range baseline tests passed without baseline regeneration.

Latest SVG sparkline primitive consolidation checkpoint: `OfficeSvgFormatting` now owns complete SVG polyline and circle element emission for `StringBuilder` renderers. Shared sparkline line-series SVG polylines and marker SVG circles use the same point-list, center/radius, and fill helpers while preserving the approved sparkline SVG attribute order plus renderer-owned scaling, marker color, and series policy. Focused Drawing formatter/sparkline tests plus Excel sparkline and premium range baseline tests passed without baseline regeneration.

Latest SVG ellipse consolidation checkpoint: `OfficeSvgFormatting` now owns complete SVG ellipse element emission for `StringBuilder` renderers. Shared Drawing ellipse shapes use the same center/radius and fill-opacity helper while preserving exporter-owned placement, paint, stroke, and transform policy. Focused Drawing formatter/exporter tests passed without baseline regeneration.

Latest SVG path-element consolidation checkpoint: `OfficeSvgFormatting` now owns complete SVG path element emission for `StringBuilder` renderers. Shared Drawing freeform path shapes and path clip definitions use the same path-data and element-emission helper while preserving exporter-owned placement, clip, paint, stroke, and transform policy. Focused Drawing formatter/exporter tests passed without baseline regeneration.

Latest SVG clip-rectangle consolidation checkpoint: `OfficeDrawingSvgExporter` clip-path rectangle and rounded-rectangle definitions now use the shared SVG rectangle element helper instead of local element assembly. Focused Drawing clip-path/exporter tests passed without baseline regeneration.

Latest SVG positioned-text consolidation checkpoint: `OfficeTextBlockRenderer` now owns positioned SVG text/tspan element emission for callers that already resolved anchor coordinates, first baseline, line height, font, and style. Shared Drawing text boxes use that helper in `OfficeDrawingSvgExporter`, with shared fill-opacity formatting for transparent text colors. Focused Drawing renderer/exporter tests passed without baseline regeneration.

Latest Excel rotated SVG text consolidation checkpoint: `OfficeTextBlockRenderer.AppendSvgTextElement` now supports positioned underline and rotation attributes. Excel's plain rotated cell-text SVG path uses the shared Drawing helper while Excel keeps rotation interpretation, clipping, style/color resolution, and approximation diagnostics. Focused Drawing helper tests plus the public Excel rotated PNG/SVG text test passed without baseline regeneration.

Latest Excel rich-text SVG segment consolidation checkpoint: `OfficeTextBlockRenderer` now owns SVG text element emission for measured rich text segments through `AppendSvgRichTextSegment`. Excel rich cell text keeps run extraction, style fallback, line layout, cursor placement, clipping, rotation grouping, and diagnostics, while segment element assembly now uses shared Drawing text formatting. Focused Drawing helper tests plus the public Excel rich-text SVG test passed without baseline regeneration.

Latest Excel comment-title SVG consolidation checkpoint: Excel comment-body title text now uses `OfficeTextBlockRenderer.AppendSvgTextElement` instead of private SVG text assembly. Excel keeps comment body placement, callout geometry, clipping, colors, source references, and approximation diagnostics; Drawing owns the reusable title text element formatting. The public comment-body object export contract now asserts the title, bold style, and start anchor in SVG alongside the existing PNG callout evidence.

Latest Visio stencil-thumbnail SVG text consolidation checkpoint: Visio stencil preview thumbnail captions now use `OfficeTextBlockRenderer.AppendSvgTextElement` instead of private SVG text assembly. Visio keeps package/gallery/thumbnail policy, while Drawing owns caption element formatting and escaping. The browser-renderable thumbnail artifact test now asserts the generated caption shape.

Latest Excel SVG root-background consolidation checkpoint: Excel range SVG export now uses `OfficeSvgFormatting.AppendRectElement` for the root background rectangle instead of private SVG rect assembly. Excel keeps canvas dimensions, viewBox policy, and selected background color; Drawing owns the reusable rectangle element formatting. The basic public Excel PNG/SVG export contract now asserts the root background shape.

Latest Visio stencil-thumbnail wrapper SVG consolidation checkpoint: Visio stencil preview thumbnail wrappers now use shared Drawing helpers for background/border rectangles, embedded preview image, and caption text. `OfficeSvgImageRenderer.AppendImage` supports adapter-supplied `preserveAspectRatio`, so Visio keeps thumbnail layout and gallery policy while Drawing owns the reusable SVG primitive emission. The browser-renderable thumbnail artifact test now asserts the wrapper rectangles, preserve-aspect image, and caption shape.

Latest Excel top/bottom conditional-formatting checkpoint: Excel image export now renders numeric top/bottom count and percent conditional-formatting rules with solid differential fills, including tied values at the cutoff. Rule snapshots carry top/bottom rank, bottom, and percent metadata; the fluent range builder can attach fill colors for top/bottom rules. Invalid or nonnumeric top/bottom rules remain source-diagnosed with `ExcelConditionalTopBottomUnsupported` instead of being silently ignored.

Latest Excel average conditional-formatting checkpoint: Excel image export now renders above-average and below-average rules with solid differential fills through the existing conditional fill pipeline, including equal-to-average variants. Sheet-level and fluent range APIs can attach average-rule fill colors, Open XML readback resolves above/below/equal/std-dev metadata, and standard-deviation average rules plus unsupported conditional-formatting families such as date/time remain source-diagnosed.

Latest Excel text conditional-formatting checkpoint: Excel image export now renders contains-text, not-contains-text, begins-with, and ends-with rules with solid differential fills through the existing conditional fill pipeline. Sheet-level and fluent range APIs can attach text-rule fill colors, Open XML readback resolves the text payload, and malformed text rules emit `ExcelConditionalTextRuleUnsupported`.

## Goal

Assess how OfficeIMO could support Excel-to-image conversion, starting with an Excel range or worksheet exported as PNG/SVG, and then extending the same pattern toward broader Office document-to-image flows. The target is no new runtime dependencies. `C:\Support\GitHub\ChartForgeX` is a useful reference implementation to borrow from, but should not become a package dependency.

## Short Answer

This is feasible, and the current branch proves the first slice without adding runtime dependencies. The feature should still not become a generic "render any Office file" promise until Excel image export reaches a premium fidelity bar and the remaining Visio renderer brain is consolidated into `OfficeIMO.Drawing`.

```csharp
byte[] png = sheet.Range("A1:D12").ToPng(new ExcelImageExportOptions { Scale = 2 });
sheet.SaveRangeAsPng("A1:D12", "summary.png");
```

The continuing safe path is to keep Excel as a thin document adapter over neutral visual snapshots, keep all pixel/path/image primitives in `OfficeIMO.Drawing`, and reuse the worksheet extraction and layout planning already built for `OfficeIMO.Excel.Pdf` where that removes duplication. Avoid routing through PDF for the primary API: the current PDF raster baseline tests use external Poppler tooling, which conflicts with the no-dependency requirement for a product feature.

## What We Have

### Excel worksheet model

`OfficeIMO.Excel` already exposes most of the raw data needed to describe a visible range:

- Range reading through `ExcelSheetReader.ReadRange(...)`, streaming variants, and used-range detection.
- Cell display text through `ExcelSheet.TryGetCellText(...)`.
- Cell style snapshots through `ExcelCell.GetStyle()` and `ExcelSheet.GetCellStyle(...)`.
- Row and column metadata through `GetRowDefinitions()` and column definition APIs.
- Merged ranges through `GetMergedRanges()`.
- Worksheet images through `ExcelSheet.Images`, `GetImage(...)`, and `AddImage(...)`.
- Chart snapshots through `ExcelChart.TryGetSnapshot(...)`.
- Shared drawing primitives through `OfficeIMO.Drawing`.

### Excel-to-PDF adapter

`OfficeIMO.Excel.Pdf` is the strongest existing implementation source. It already translates workbook content into a planned export model:

- `ExcelPdfSaveOptions` covers selected sheets, print areas, page setup, headers/footers, cell styles, hyperlinks, images, charts, merged cells, column widths, row heights, hidden rows/columns, and diagnostics.
- `BuildWorksheetExportPlans(...)` chooses sheets/ranges and builds `WorksheetPdfExportPlan`.
- `ReadRangeExportData(...)` returns values, styles, hyperlinks, cell references, merge layout, column widths, and row heights.
- `WorksheetImageExportData` and `WorksheetChartExportData` already model media attached to a worksheet export.
- PDF tests cover many Excel export contracts: charts, images, styles, layout, print areas, page setup, links, and hidden rows/columns.

This means the missing work is mostly "turn a worksheet visual plan into pixels", not "understand Excel from scratch".

### Shared drawing layer

`OfficeIMO.Drawing` is zero-dependency and now has:

- `OfficeDrawing` as a vector drawing canvas.
- `OfficeShape`, `OfficeDrawingShape`, `OfficeDrawingText`, colors, gradients, shadows, transforms, clipping, and chart snapshots.
- `OfficeDrawingSvgExporter.ToSvg(...)`.
- Deterministic text measurement through `OfficeTextMeasurer`.
- Shared chart rendering through `OfficeChartDrawingRenderer`.
- `OfficeRasterImage`, `OfficeRasterCanvas`, `OfficeRasterRenderTarget`, PNG read/write, alpha blending, antialiasing, downsample resolve, image compositing, polygon/contour fill helpers, line/polyline/dashed stroke helpers, dashed ellipse stroke approximation, solid elliptical arcs, rotated ellipse fill/stroke, rotated/scaled image projection, raster path flattening for Bezier path commands, anchored text-line rendering, fallback glyph rendering, rectangular clipping scopes, and shared text measurement.
- Shared hatch-pattern primitives for raster and SVG output, with Excel-specific pattern semantics kept in the Excel adapter.
- Shared sparkline primitives for raster and SVG line, column, and win/loss output, with Excel-specific source extraction and diagnostics kept in the Excel adapter.
- Shared resolved data-bar primitives for raster and SVG output, with Excel-specific conditional-formatting rule evaluation kept in the Excel adapter.
- `OfficePathCommand` plus `OfficeSvgFormatting.FormatPathData(...)` / `AppendPathData(...)` for shared SVG `M`/`L`/`Q`/`C`/`Z` path serialization.
- `OfficeStrokeDashStyleMapper` for shared Visio line-pattern and Office preset-dash normalization without taking an OpenXML dependency.

The remaining issue is not absence of a raster stack; it is fidelity depth and completing consolidation so Excel, Visio, and later document families do not maintain separate renderer brains.

### ChartForgeX reference material

ChartForgeX is explicitly dependency-free and has mature raster infrastructure that maps well to OfficeIMO's needs:

- `RgbaImage` and `RgbaCanvas`.
- `PngWriter`, plus GIF/JPEG/BMP/PPM/TIFF encoders.
- Text drawing and measuring with built-in tiny font fallback plus optional TrueType loading.
- SVG/PNG parity patterns for charts, tables, visual blocks, and canvases.
- `TableArtifact` and `ChartTable` preview rendering, useful as a reference for simple table-like image output.

The useful borrowing target is the low-level raster stack and renderer discipline, not ChartForgeX chart/domain models.

### Existing PDF raster tests

`OfficeIMO.Tests` has PDF raster visual baseline tests, including native Excel-to-PDF baselines. These tests rasterize generated PDFs with Poppler `pdftoppm`, compare PNG pixels through the shared Drawing-backed visual-baseline helper, and store approved PNGs. That is useful QA evidence, but it is test infrastructure only and depends on an external executable.

## What Is Missing For Premium Excel Export

The first slice proves the architecture. It is not premium yet. The big missing groups are:

- Text/layout: basic wrapping, explicit/default vertical alignment, styled font sizes, basic shrink-to-fit, basic numeric text rotation, basic stacked plain text, SVG and PNG cell text clipping without synthetic ellipsis, single-line, hard-break, shrink-to-fit, basic rotated rich text run rendering, and basic stacked rich text run rendering, Drawing-level text measurement caching, Drawing-level rectangular clip scopes, shared Drawing word wrapping/line measurement for Visio native PNG and SVG text, shared Drawing line layout and overflow policy for Excel plain multiline cell text, shared Drawing stacked text-block layout for Excel `TextRotation=255`, shared Drawing stacked rich-text block layout for Excel `TextRotation=255`, shared Drawing plain text-block PNG/SVG emission for non-rotated Excel cell text, shared Drawing plain text-block PNG emission for Visio text, shared Drawing SVG text/tspan emission for Visio text, shared Drawing rich-run block layout for Excel hard-break/wrapped/shrink-to-fit/rotated/stacked rich text, and `ExcelCellTextClipped`, `ExcelCellTextRotationApproximation`, `ExcelCellRichTextLayoutApproximation`, and `ExcelCellFontFamilyFallback` diagnostics now exist. Premium still needs Excel-exact text spill into adjacent empty cells, Excel-exact clipped/rotated/stacked rich text behavior beyond the current run-preserving approximation, Excel-exact rotated/stacked baseline semantics, proper baseline metrics, broader font matrices, closer font-aware measurement parity, and deeper clipping parity for images, charts, shapes, and every layered object path.
- Styles: first-pass percent/number/date display text now uses a shared helper also used by autofit, and that helper now preserves custom literal affixes, escaped literal characters, and positive/negative/zero format sections for image snapshots and SVG output; direct/theme/indexed color resolution with tint/shade support is shared by `GetStyle()`, inspection snapshots, and image rendering, conditional color scales/data bars render through the neutral Excel image snapshot, bounded numeric cell-is plus simple comparison formula differential fills now honor priority and stop-if-true behavior, top/bottom count and percent rules, duplicate/unique-value rules, above/below-average rules, and contains/not-contains/begins-with/ends-with text rules render solid differential fills with source diagnostics for unsupported variants, Excel image output maps common solid, hairline, medium, thick, dashed, dotted, dash-dot, dash-dot-dot, double-line, and diagonal border styles through shared Drawing stroke primitives, non-solid pattern fills now render as deterministic hatch approximations with source diagnostics, and simple two-stop linear gradient fills now render in PNG/SVG through shared Drawing gradient primitives. Premium still needs broader date/currency/custom/scientific/conditional number-format parity, full conditional-formatting parity for icon sets, date/time, and standard-deviation average variants, richer differential formats, exact complex/path/multi-stop gradient fills, more exact pattern fill density/parity, style inheritance edge cases, and a broader baseline matrix.
- Images: embedded PNG rasterization works for PNG output; SVG output can embed detected PNG, JPEG, GIF, and SVG worksheet images; range exports include and clip worksheet images that visually overlap the selection even when the anchor cell is outside it; basic two-cell anchored worksheet pictures now size from their from/to marker geometry; authored `a:srcRect` crop rectangles, basic rotation, and horizontal/vertical flips now render in PNG/SVG for worksheet pictures, including combined crop-plus-flip-plus-rotation output through one shared Drawing image projector; shared Drawing now also owns reusable image target-box, fit, and source-crop render-plan math used by PDF image placement; supported mixed shape/image/chart overlays now paint in source drawing order; unsupported PNG rasterization and SVG embedding still emit stable image-source diagnostics. Premium still needs full transform/effects parity beyond basic crop/flip/rotation, transparency edge cases, hidden-row/column clipping parity, EMF/WMF handling, first-party decoding for more raster formats, and stronger diagnostics for every unsupported image effect.
- Charts: current chart export is still an approximation through shared chart snapshots, but the first authored chart layout/style bridge now exists. Chart and plot area solid fill/outline/width/preset dash, legend layout, title overlay, axis titles, chart-level data-label layout, simple authored series colors, line widths, and preset dashes, simple point fills, simple marker fill colors, marker visibility, simple marker size, simple marker solid outline color/width, simple circle/square/diamond/triangle/dash/dot/plus/X/star marker shapes, simple category/value major-gridline color/visibility/width/preset dash, simple category/value minor-gridline color/visibility/width/preset dash, simple category/value axis-line color/visibility/width/preset dash, simple category/value major tick marks, simple category/value minor tick marks, category/value axis label visibility when Excel tick-label position is `none`, simple high/low/next-to category/value tick-label placement, simple maximum-crossing horizontal category-axis and vertical value-axis placement, simple category/date-axis reverse-order rendering, simple value-axis number formats for vertical and horizontal bar orientations, default value-axis major-tick label density, simple value-axis display-unit scaling and labels, simple linear value-axis min/max/major/minor-unit scaling, simple chart title text color/font-family/font-size/bold/italic, simple legend text color, simple data-label text color, simple axis-label text color, simple axis-title text-color override, simple legend/data-label/axis-label font sizes, simple axis-title font size, simple legend/data-label/axis-label font families, simple legend/data-label/axis-label bold/italic buckets, and simple axis-title font-family/bold/italic overrides flow to shared Drawing; `ExcelChartTrendlineUnsupported`, `ExcelChartDataLabelPointOverridesApproximated`, `ExcelChartDataLabelLeaderLinesUnsupported`, `ExcelChartAreaStyleApproximation`, `ExcelChartGridlineStyleApproximation` for complex gridline effects, `ExcelChartAxisStyleApproximation`, `ExcelChartAxisTickLabelPositionApproximation`, `ExcelChartAxisMinorTickMarkPlacementApproximation` for remaining approximate minor tick-mark placement, `ExcelChartAxisCrossingApproximation`, `ExcelChartAxisScaleApproximation`, `ExcelChartAxisNumberFormatApproximation`, `ExcelChartCategoryAxisNumberFormatUnsupported`, `ExcelChartTextStyleApproximation`, `ExcelChartFontFamilyFallback`, `ExcelChartSeriesStyleApproximation`, chart-kind approximation, and unsupported-kind diagnostics make the remaining gaps visible. Premium still needs picture markers and richer marker outline effects, custom/richer series dash/effect styling, richer point-level overrides, trendline rendering, leader lines, point-level label overrides, richer axes/tick formatting beyond simple value-axis numeric formats and simple high/low/next-to/none tick-label placement, Excel-exact display-unit placement/typography, category/date-axis tick formatting, log/value-axis-reverse-order/non-value-axis-unit/non-default cross-between axis rendering and crossing-at values, remaining Excel-exact minor-gridline/tick placement edge cases, custom dash/effect parity beyond preset gridline and axis lines, chart/plot area effects beyond simple solid RGB fill/outline/width/preset dash, richer chart title typography/effects beyond simple font-family/font-size/bold/italic, richer/per-element chart rich text runs beyond the supported shared buckets, Excel-exact chart geometry, and diagnostics for every impossible-parity chart feature.
- Worksheet/page behavior: hidden rows/columns now have an explicit range-export contract: omitted by default, included with `IncludeHidden`, and reported with source diagnostics when omitted or when anchored images/charts are skipped because their anchor is hidden. Worksheet image export can honor a configured print area through `ExcelWorksheetImageExportOptions.UsePrintArea`; explicit `Range` still wins, missing print areas report `ExcelPrintAreaMissing`, and multi-area print areas can be returned as separate images through the multi-output export path. Manual page breaks can split worksheet image exports, repeated print-title rows/columns are composed over each sliced page, and supported text-only first/even/odd header/footer chrome renders through shared Drawing text/image composition. Page-sliced output now applies orientation, margins, manual scale, supported paper-size geometry, and bounded one-page fit-to-width/fit-to-height scaling onto a physical page canvas through shared image composition, with missing or unknown paper sizes and unsupported multi-page fit requests explicitly diagnosed. Excel image export and first-party Excel PDF now share worksheet paper-size resolution through `ExcelPageSetupGeometry`, and Excel PDF honors supported worksheet paper-size codes when no explicit PDF page size is supplied. Premium still needs automatic multi-page fit pagination, broader paper-size coverage, large-sheet tiling beyond manual page breaks, Excel-exact header/footer image support, and deeper hidden-row/column parity for page-oriented worksheet export.
- Objects: worksheet hyperlink visual hints now flow through the Excel visual snapshot and can be disabled with `ShowHyperlinkHints`; range hyperlinks are normalized for inspection and image snapshots; comments/notes and threaded comments are detected through shared metadata resolution, render visible cell indicators in PNG/SVG, can opt into first-pass dependency-free body callouts with `ShowCommentBodies`, paint callout body text through shared Drawing text-block emission, and report either unsupported-body diagnostics or rendered-body approximation diagnostics depending on that option; worksheet drawing objects are classified through one shared resolver, simple rectangle/rounded-rectangle shapes/text boxes and supported preset shapes with solid RGB fill/outline render through shared Drawing, authored shape rotation is carried through the shared transform path, and supported shapes/images/opt-in comment bodies/charts now render through one ordered overlay stream. Premium still needs Excel-exact comment/note/threaded-comment popover geometry, visibility state, connectors, threading metadata, stacking, richer shapes/text boxes/connectors, theme/system/transformed colors, Excel-exact rotated shape text metrics, grouping, richer hyperlink affordances, deeper layered-object clipping, and more polished preset geometry parity.
- Diagnostics: unsupported image formats, SVG embedding gaps, hidden row/column omissions, hidden-anchored image/chart omissions, print-area fallback cases, text clipping, explicit cell/header-footer/chart font-family fallback, rich-text rotation fallback, rotation and stacked-text approximations, rendered sparkline approximations, unsupported or unresolved sparkline data cases, unsupported conditional icon sets, unsupported conditional rule types, unsupported conditional cell-is/formula shapes, unsupported conditional differential-format effects, unsupported comment/note/threaded-comment bodies, opt-in comment/note/threaded-comment body approximations, unsupported worksheet drawing geometry/fill/outline/rotation/connectors/groups/non-chart frames, and chart approximations including richer chart text styling now have stable diagnostics. Premium still needs the same treatment for every remaining unsupported or approximated feature, especially richer conditional-formatting subfeatures, richer chart title typography/effects beyond simple font-family/font-size/bold/italic, per-element chart text typography, and object/page-layout effects, with stable codes and source references.
- Visual QA: Excel image PNG/SVG visual baseline gates now validate a curated range with styled cells, merged title, percent display text, wrapping, clipped text without synthetic ellipsis, vertical alignment, single-line rich text, a comment indicator, embedded image, and chart snapshot; a dedicated chart-axis label range with intermediate value-axis tick labels aligned to major gridlines; a dedicated rich-text range with single-line, hard-break, wrapped, shrink-to-fit, clipped rich text without synthetic ellipsis, and basic rotated rich text; a dedicated stacked-text range with upright `TextRotation=255` plain and rich output, run-level rich styling, and no old unsupported or rich-layout approximation diagnostic; a dedicated pattern-fill range with horizontal, vertical, grid, diagonal, trellis, and dotted hatch output; a conditional-formatting range with heat-map fills, positive and negative data bars, a rule-driven differential fill column, and unsupported icon-set diagnostics; a sparkline range with line, column, and win/loss output; a drawing-object range with a visually reviewed simple shape/text-box rendered through shared Drawing; a clipped-image range where the exported selection cuts through an image anchored outside the range; a two-cell image range that proves worksheet picture sizing from marker geometry; a cropped-image range that proves `a:srcRect` output; a visually reviewed rotated-image range that proves authored picture rotation without obscuring surrounding text; and a visually reviewed transformed-image range that combines source crop, horizontal flip, and rotation. Focused text/contracts also validate rotated PNG text clipping, SVG single-line rich-text run styling, hard-break rich-text PNG/SVG preservation, shrink-to-fit rich-text PNG/SVG preservation, plain/rich overflowing cell SVG output without ellipsis plus clipping diagnostics, basic rotated rich-text run preservation with diagnostics, stacked plain and rich text PNG/SVG rendering through shared Drawing layout, source-filtered rendered comment/threaded-comment indicators with decoded PNG pixels and SVG polygons, opt-in rendered comment body callouts with drawing-layer placement, decoded PNG pixels, and SVG text/color assertions, supported drawing-object SVG/PNG output and decoded fill pixels, mixed shape/image z-order with decoded PNG pixels and SVG order assertions, range-clipped overlapping image decoded PNG pixels and SVG clip paths, two-cell image marker sizing with decoded PNG pixels, cropped worksheet image decoded PNG pixels and SVG clipping, authored worksheet image rotation with decoded PNG pixels outside the unrotated rectangle and SVG rotation transforms, combined authored crop/flip/rotation with decoded PNG pixels and SVG transforms, source-filtered worksheet drawing-shape diagnostics, page-sliced print-title composition, rendered header/footer text chrome, page-setup PNG/SVG canvas dimensions plus margin placement, supported A4/Legal paper-size dimensions, unknown paper-size fallback diagnostics, bounded one-page fit-to-width/fit-to-height scaling, shared Drawing point conversion, Excel PDF worksheet paper-size output, rendered line/column/win-loss sparkline SVG structure, decoded PNG sparkline pixels, and cross-sheet sparkline diagnostics. Shared raster baseline comparison support now covers Excel, Visio, and PDF tests, and Visio premium native baselines are split per scenario so one visual regression can be reviewed without rebuilding the whole gallery. Premium still needs a broader Excel baseline matrix for Excel-exact text spill/overflow behavior, rotated/stacked rich text, deeper chart fidelity, richer image anchors/effects, automatic multi-page fit pagination and tiling geometry, broader paper-size coverage, Excel-exact object bodies and richer drawings, Excel-exact sparkline parity, and impossible-parity diagnostics.

Sparkline status: authored sparklines are no longer a hidden omission for image export. They are detected from worksheet extension metadata, filtered to the exported visible range, rendered for same-sheet numeric line/column/win-loss cases, and diagnosed with stable source codes when the renderer approximates or cannot resolve the data. Approved PNG/SVG baselines now cover the initial sparkline variants. Premium still needs Excel-exact hidden/empty data behavior, group-level axis/date/scaling parity, and broader baselines for those deeper variants.

### Excel Image Visual QA Gate

Use `Build/Test-ExcelImageVisualGate.ps1` as the local review gate for Excel image PNG/SVG work. It intentionally splits the visual checks instead of using one silent mega-filter. The default `Full` suite remains the release/review gate:

```powershell
Build/Test-ExcelImageVisualGate.ps1 -NoRestore -NoBuild
```

The full suite runs:

- generated Excel image output vs approved PNG/SVG baselines
- approved baseline renderability/nonblank checks
- `DrawingArchitectureTests` to keep Excel/Visio/PDF image paths routed through the shared dependency-free `OfficeIMO.Drawing` brain

Use the smoke suite for local iteration when a change needs fast visual evidence before the full gate:

```powershell
Build/Test-ExcelImageVisualGate.ps1 -Suite Smoke -NoRestore -NoBuild
```

The smoke suite covers representative premium range, rich text, header/footer image, chart-axis, page-layout, conditional-formatting, drawing-object, and transformed-image baselines plus the same shared Drawing architecture guard. Use `-Suite Architecture` when only the shared dependency-free rendering-owner guard needs to run. Use `-UpdateBaselines` only after inspecting the generated artifacts and deciding the visual change is intentional. The generated-output half of the full suite is the expensive lane; on the current branch it has taken roughly 8-10 minutes with normal console progress, while the smoke suite takes roughly 3-4 minutes and the architecture suite stays near a few seconds.

## What Was Missing At Assessment Start

### 1. A product raster layer

OfficeIMO had SVG export for `OfficeDrawing`, but no first-party raster layer equivalent to ChartForgeX's `RgbaCanvas` and `PngWriter`. The current branch adds the initial shared raster layer; the remaining work is hardening and completing feature coverage.

Minimum reusable pieces:

- `OfficeRasterImage` or `OfficeRgbaImage`.
- `OfficeRasterCanvas`.
- `OfficePngWriter`.
- Basic drawing operations: fill/stroke rect, rounded rect, line, polygon/path enough for borders/icons/data bars, text, image compositing.
- A small public raster output model: bytes, width, height, MIME type, diagnostics.

Best home: `OfficeIMO.Drawing`, because Word, PowerPoint, Excel, PDF, Markdown, RTF, Visio, and future adapters can all reuse it without introducing a new dependency.

### 2. A public worksheet visual snapshot

`OfficeIMO.Excel.Pdf` has excellent private planning models, but they are PDF-adapter internals. Excel-to-image should not copy them. We should promote a neutral snapshot layer, for example:

- `ExcelRangeVisualSnapshot`
- `ExcelRangeVisualCell`
- `ExcelRangeVisualStyle`
- `ExcelRangeVisualMerge`
- `ExcelRangeVisualImage`
- `ExcelRangeVisualChart`
- `ExcelRangeVisualOptions`
- `ExcelRangeVisualDiagnostics`

This layer belongs in `OfficeIMO.Excel` if it is pure workbook understanding, or in a thin sibling adapter if it requires image-only options. The snapshot should be independent of PDF and raster outputs.

### 3. A range renderer

Need a renderer that maps the neutral range snapshot to:

- SVG, likely via `OfficeDrawing` or a dedicated SVG table renderer.
- PNG, via the new OfficeIMO raster layer.

The renderer should handle:

- Cell backgrounds.
- Gridlines and explicit borders.
- Text alignment, font style, wrapping, and clipping.
- Number/date display text as already produced by Excel read/display helpers.
- Merged cells.
- Row heights and column widths.
- Hidden rows/columns.
- Conditional fills, data bars, and simple icon sets.
- Images anchored inside the exported range.
- Chart snapshots inside the exported range, using `OfficeChartDrawingRenderer`.

### 4. Output APIs and diagnostics

Need a small, friendly API surface:

```csharp
public sealed class ExcelImageExportOptions {
    public string? SheetName { get; set; }
    public string? Range { get; set; }
    public int Scale { get; set; } = 2;
    public bool ShowGridlines { get; set; } = true;
    public bool RespectHiddenRowsAndColumns { get; set; } = true;
    public bool IncludeImages { get; set; } = true;
    public bool IncludeCharts { get; set; } = true;
    public string EmptyCellText { get; set; } = string.Empty;
}

public sealed class ExcelImageExportResult {
    public byte[] Bytes { get; }
    public int WidthPixels { get; }
    public int HeightPixels { get; }
    public string ContentType { get; }
    public IReadOnlyList<ExcelImageExportWarning> Warnings { get; }
}
```

Candidate extension methods:

```csharp
sheet.ToPng("A1:D12", options);
sheet.SaveAsPng("A1:D12", path, options);
sheet.ToSvg("A1:D12", options);
document.SaveSheetAsPng("Summary", path, options);
document.SaveAsImages(directory, options);
```

### 5. Fidelity rules

Need explicit boundaries. Excel range-to-image should be a polished OfficeIMO-rendered visual approximation, not an Excel screenshot clone and not a rough debug rendering.

Expected first milestone:

- Good for generated OfficeIMO workbooks and common OpenXML workbooks.
- Deterministic cross-platform output.
- No Excel, LibreOffice, browser, GDI, System.Drawing, Skia, ImageSharp, Playwright, or Poppler requirement.
- Visually credible enough for reports, documentation, previews, and user-facing exports.
- Human visual review or approved visual baselines for renderer changes.

Known first-wave gaps:

- Exact Excel font metrics.
- Formula recalculation when cached values are missing.
- Complex style interactions beyond direct/theme/indexed colors and first-pass tint/shade resolution.
- Advanced conditional formatting.
- Rich text runs inside one cell.
- Excel-exact rotated and stacked text.
- Precise drawing z-order and two-cell anchor clipping.
- Pictures with crop/transform effects.
- Excel-native chart fidelity beyond current chart snapshot support.
- Frozen panes, selections, rendered comments/notes/threaded comments, slicers, pivots, Excel-exact sparklines, richer shapes/text boxes/connectors, SmartArt, OLE objects.

Quality bar before treating this as product-ready:

- Text and chart labels must be clean, spaced, and legible.
- Cell padding, row heights, column widths, borders, and merged cells must feel Office-like.
- Embedded images must keep transparency/aspect and land in the expected visual position.
- Charts must look intentional: title, axes, gridlines, legend, and series geometry should not read as hand-authored placeholders.
- Visual test artifacts must be opened during review, not only checked for valid bytes.

## Recommended Architecture

### Layer 1: OfficeIMO.Drawing.Raster

Add raster primitives to `OfficeIMO.Drawing` by adapting the dependency-free ChartForgeX raster stack:

- Start with RGBA canvas and PNG writer only.
- Keep encoders for JPEG/GIF/BMP/TIFF out of the MVP unless there is a concrete user need.
- Map `OfficeColor` instead of ChartForgeX `ChartColor`.
- Keep all code dependency-free and compatible with `netstandard2.0`, `net8.0`, `net10.0`, and Windows `net472`.
- Add tests that verify PNG signatures, dimensions, transparent pixels, and simple drawing output.

### Layer 2: OfficeIMO.Drawing renderers

Add:

- `OfficeDrawingRasterRenderer` for `OfficeDrawing -> OfficeRasterImage`.
- `OfficeDrawingPngExporter.ToPng(...)`.

This makes existing chart snapshots and visual drawings exportable to SVG and PNG from the shared core.

### Layer 3: OfficeIMO.Excel visual snapshot

Extract and generalize the private models from `OfficeIMO.Excel.Pdf`:

- Move workbook/range extraction that is not PDF-specific into `OfficeIMO.Excel`.
- Keep `OfficeIMO.Excel.Pdf` as a thin adapter over the shared snapshot.
- Preserve its current tests while adding image-specific tests.

### Layer 4: OfficeIMO.Excel image adapter

Add a small Excel-owned adapter surface that owns only image-oriented options and extension methods. It should call the shared snapshot extractor and shared renderer.

First implementation package:

- `OfficeIMO.Excel`

First implementation namespace:

- `OfficeIMO.Excel`

Do not create `OfficeIMO.Excel.Image` for the first implementation. Excel already references `OfficeIMO.Drawing`, so a new package adds release and discovery friction without reducing runtime dependencies. Reconsider a separate package only if the image surface becomes large enough to justify optional install size or release-cadence separation. Internally, the reusable snapshot still belongs in `OfficeIMO.Excel` so PDF and image do not diverge.

### Layer 5: All-to-image

After Excel range/image works, the same raster core can support:

- `OfficeDrawing -> PNG/SVG`.
- `OfficeChartSnapshot -> PNG/SVG`.
- PowerPoint slide -> PNG, by reusing the existing PowerPoint-to-PDF/shape rendering knowledge and `OfficeIMO.Drawing`.
- Word page/section -> PNG, probably later because pagination fidelity is harder.
- PDF page -> PNG only if OfficeIMO gains a real PDF content rasterizer. Current Poppler tests should remain QA-only unless external dependencies are explicitly allowed.

## Suggested MVP Scope

1. Add `OfficeIMO.Drawing` raster primitives and PNG writer.
2. Add `OfficeDrawing` to PNG support.
3. Add `ExcelRangeVisualSnapshot` extraction for values, styles, merges, row/column sizes, hidden rows/columns, images, and chart snapshots.
4. Add `ExcelRangeImageRenderer` with cell fills, borders, text, merges, row/column sizing, simple conditional fills, images, and chart snapshots.
5. Add extension methods for `ExcelSheet.ToPng(range)` and `SaveRangeAsPng(...)`.
6. Add tests:
   - generated workbook range to PNG has valid PNG dimensions and non-empty visible pixels;
   - fill/border/text smoke test;
   - merged cells smoke test;
   - hidden row/column omission;
   - anchored image appears;
   - chart snapshot path emits nonblank image;
   - SVG and PNG dimensions agree.

## Risk Notes

- Do not make `OfficeIMO.Excel.Pdf` internals the public image API. Promote neutral snapshot contracts first.
- Do not depend on ChartForgeX at runtime. Borrow/adapt the small raster primitives if license and ownership are acceptable.
- Do not use Poppler, Excel automation, LibreOffice, browser screenshots, `System.Drawing.Common`, SkiaSharp, ImageSharp, or QuestPDF for the core path.
- Do not promise exact Excel screenshots. Promise deterministic OfficeIMO rendering with explicit diagnostics and documented unsupported features.
- Keep the first implementation range-focused. Whole workbook and all-document image export can be built as orchestration over page/sheet/range images.

## Bottom Line

OfficeIMO is closer than it first looks. The workbook understanding, style extraction, chart snapshots, images, and PDF export planning already exist. The missing reusable piece is a dependency-free raster layer plus a neutral Excel visual snapshot/renderer. Borrowing ChartForgeX's raster approach into `OfficeIMO.Drawing` is the right way to keep the implementation first-party, reusable, and dependency-free.

Latest shared-text checkpoint: `OfficeIMO.Drawing` now owns a reusable `OfficeTextBlockRenderPlan` for fitted center-based text placement, left/top rectangle cell placement, background bounds, and rotated background corners, plus shared SVG rich-text block emission for measured rich-run output. Visio SVG/PNG text adapters, Excel PNG/SVG plain cell text, and Excel SVG rich text now consume Drawing helpers instead of carrying separate placement math, giving premium Excel image-export work another non-Excel proof point for the shared Drawing engine.
