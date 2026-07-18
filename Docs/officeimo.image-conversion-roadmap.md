# OfficeIMO Image Conversion Roadmap

Date: 2026-06-23

Implementation status: phases 1-5 have an initial vertical implementation for the Excel goal. Excel ranges, worksheets, and workbooks export to PNG/SVG through a shared dependency-free `OfficeIMO.Drawing` raster/PNG foundation and Excel visual snapshots. The renderer paints cells, fills, styled borders, diagonal borders, gridlines, display text, single-line, hard-break, shrink-to-fit, basic rotated rich text runs, and basic stacked rich text runs, merged-cell coverage, hyperlink visual hints, Excel-style comment indicators and opt-in comment body callouts, supported simple worksheet drawing shapes/text boxes, supported rotated DrawingML preset shapes, embedded PNG worksheet images, range-clipped worksheet images that visually overlap the selected range even when anchored just outside it, SVG-embeddable worksheet image formats, first-pass conditional color scales/data bars, bounded numeric cell-is/formula differential fills, supported chart snapshots, repeated print-title rows/columns for manual-page-sliced multi-output worksheet exports, and first/even/odd text header/footer chrome with supported page/sheet/workbook file/date/time fields, basic font-family/bold/italic/underline/color/font-size/strikethrough formatting approximation, plus clipped and ellipsized left/center/right zones for manual-page-sliced image exports. A first visual-fidelity pass added proper source-over alpha blending, antialiased text/shape edges, interpolated image scaling, raster text alignment/style mapping, shared styled-line rendering for solid, dashed, dotted, dash-dot, dash-dot-dot, and double-line Excel border output, shared Office/Visio dash vocabulary normalization, and normalized worksheet hyperlink resolution shared by inspection and image snapshots. The current Excel fidelity work adds wrapped cell text, explicit vertical text alignment plus Excel-like bottom alignment when the source style is unset, styled font sizes, basic shrink-to-fit, basic numeric text rotation, per-cell SVG clipping, PNG text clipping through shared Drawing raster clip scopes, shared plain text-block raster/SVG rendering for non-rotated Excel cell text plus Visio PNG/SVG text, shared rich text block layout and shared raster rich-text block rendering for hard-break, wrapped, shrink-to-fit, basic rotated runs, and basic stacked runs, direct/theme/indexed cell color resolution with tint/shade support for fills, fonts, borders, and sheet tabs, dependency-free hatch approximations for Excel pattern fills, simple two-stop linear gradient fills through shared Drawing raster/SVG primitives, shared PNG/SVG image layer composition for page-level print-title and header/footer assembly, explicit hidden row/column omission behavior with `IncludeHidden`, rendered source-referenced comment/note and threaded-comment cell indicators plus opt-in approximation callouts backed by a shared worksheet comment resolver, shared worksheet drawing-object classification for PDF preflight and image rendering/diagnostics, shared shape-transform rendering for Drawing raster output, and a source-order drawing layer that renders supported shapes, images, charts, and opt-in comment bodies through one mixed overlay stream instead of separate renderer brains. Stable `ExcelCellTextClipped`, `ExcelCellTextRotationApproximation`, `ExcelCellStackedTextRotationUnsupported`, `ExcelCellRichTextLayoutApproximation`, `ExcelFillPatternApproximation`, `ExcelFillGradientUnsupported`, `ExcelConditionalIconSetUnsupported`, `ExcelConditionalRuleUnsupported`, `ExcelConditionalCellIsUnsupported`, `ExcelConditionalFormulaUnsupported`, `ExcelHiddenRowsOmitted`, `ExcelHiddenColumnsOmitted`, `ExcelImageAnchorHidden`, `ExcelChartAnchorHidden`, `ExcelCellCommentUnsupported`, `ExcelCellCommentBodyApproximation`, `ExcelThreadedCommentUnsupported`, `ExcelThreadedCommentBodyApproximation`, `ExcelDrawingShapeUnsupported`, `ExcelDrawingShapeTextRotationApproximation`, `ExcelDrawingShapeTextAutoFitUnsupported`, `ExcelDrawingShapeTextVerticalOrientationUnsupported`, and `ExcelHeaderFooterFormattingApproximation` diagnostics now exist, along with shared number-format display text for the Excel image and autofit paths, including custom literal affixes, escaped literal characters, and positive/negative/zero section selection. The approved Excel image PNG/SVG visual baselines now cover merged title text, styled cells, percent display text, wrapping, clipping, vertical alignment, single-line rich text, an Excel-style comment indicator, a supported drawing-object shape, an embedded PNG image, a range-clipped overlapping image, a chart snapshot, conditional heat-map fills, positive/negative data bars, bounded cell-is differential fills, and unsupported icon-set diagnostics; focused contract tests cover default bottom vertical alignment, custom number-format literal/section display text, shared plain text-block raster/SVG rendering, shared raster rich-text block rendering, shared SVG `text`/`tspan` writer output, shared transformed-shape raster output, hyperlink hint SVG output, range hyperlink expansion, rotated PNG text clipping, single-line, hard-break, shrink-to-fit, basic rotated rich text SVG/PNG styling, and basic stacked rich text SVG/PNG styling, OpenXML pattern fill SVG/PNG rendering with diagnostics, simple linear gradient SVG/PNG rendering plus unsupported-gradient diagnostics for unresolved cases, source-filtered rendered comment/threaded-comment indicators with unsupported body diagnostics, opt-in comment body callouts with approximation diagnostics, drawing-layer placement, anchored SVG pointers, and decoded PNG pixels, supported rounded-rectangle and preset drawing-object output through shared Drawing, rotated preset drawing-object SVG/PNG output, mixed shape/image z-order in both directions with decoded PNG pixels and SVG order assertions, range-clipped overlapping worksheet images, supported font-family/color/font-size/strikethrough header/footer SVG output, and source-filtered worksheet drawing-shape diagnostics. The shared raster stack now also has first migrated primitives from the Visio renderer needs: solid and styled polyline strokes, solid elliptical arcs, polygon strokes, even-odd multi-contour fills, shape transforms for rectangle/rounded-rectangle/ellipse/polygon/path rendering, rotated/scaled image drawing, rotated ellipse fill/stroke, anchored text-line rendering, shared plain text-block rendering, fallback glyph rendering, cached text measurement, rectangular clipping scopes, linear-gradient rectangle fills, supersampled render-target storage, alpha blending, adaptive coverage sampling for supersampled render targets, downsample resolve, and shared Visio line-pattern plus Office preset-dash mapping. Excel, Visio, and PDF raster visual-baseline tests now share one Drawing-backed PNG decode/encode/diff helper. Unsupported image rasterization formats, SVG image embedding limits, hidden rows/columns omitted from the selected visual range, hidden-anchored worksheet images/charts, stacked text rotation, numeric/rich text-rotation approximation, rotated drawing-object text approximation, unsupported shape-text autofit and complex vertical orientation, approximated pattern fills, unsupported complex/path/multi-stop gradient fills, unsupported conditional icon sets, unsupported conditional rule types, unsupported conditional cell-is/formula shapes, richer header/footer font style variants and images, default-disabled comment/note/threaded-comment bodies, opt-in comment body approximations, unsupported or richer worksheet drawing shapes/text boxes/connectors, and approximated chart kinds are surfaced through diagnostics instead of being hidden.

Consolidation status: the current branch proves the shared `OfficeIMO.Drawing` direction, but OfficeIMO still has more than one rendering engine. `OfficeIMO.Visio` has a mature private PNG renderer whose PNG decoder, PNG encoder, supersampled pixel storage, source-over alpha blending, downsample resolve, polygon fills, even-odd contour fills, polyline/dashed stroke drawing, dashed ellipse stroke approximation, solid elliptical arcs, rotated ellipse fill/stroke, rotated/scaled image projection, anchored text-line rendering, fallback glyph rendering, text measuring, word wrapping/line measurement, single-line font fitting, bounded text-block fit math, bounded visible-line clipping, and reusable text placement math have been moved onto shared Drawing primitives. Its remaining private `RasterCanvas` is now a Visio-specific geometry adapter over shared raster storage and shared canvas operations, not the final shared engine shape. Visio SVG text export now also uses the shared Drawing text layout primitive for line construction, max-line measurement, bounded font fit, and text placement, so Visio no longer has separate PNG and SVG text wrapping, block-fit, or placement implementations. Visio save-time resize-to-text geometry now also routes line construction and long-word breaking through the shared Drawing text layout engine while Visio keeps document margins, padding, and inch/pixel policy. Excel plain cell text layout now consumes the same shared Drawing text layout primitive for wrapped/multiline line construction, trim-to-width, shrink-to-fit font sizing, bounded text-block orchestration, visible-height clipping, and clipped-state reporting, and Excel rich cell text now consumes shared Drawing rich-run block layout for hard-break, wrapped, and shrink-to-fit runs while keeping Excel-specific policy, vertical-alignment mapping, rotation fallback, and diagnostics in the Excel adapter. PDF has separate image/compression code because PDF streams have different writer contracts, but PDF/Visio/Excel visual-baseline PNG comparison now uses shared test support. Excel comment/threaded-comment metadata resolution is now shared by inspection, feature reporting, and image diagnostics instead of living as an inspection-only helper. Excel worksheet drawing-object classification is now shared by PDF preflight and image export, with supported simple shapes routed through shared Drawing and unsupported object variants diagnosed from the same resolver. Before premium Excel work grows much further, remaining reusable Visio raster behavior must move into `OfficeIMO.Drawing`, and Visio must be migrated to consume the shared engine without losing its existing premium visual baselines. The native Visio premium PNG baselines were refreshed after visual review for the shared renderer output, and the native premium baseline gate now runs as per-scenario theories over the shared Drawing-backed comparison helper so failures point at one approved visual at a time.

Latest consolidation checkpoint: `OfficeIMO.Drawing` now owns `OfficeTextLayoutEngine`, measured `OfficeTextLine` output, `OfficeTextBlockLayout`, `OfficeTextVerticalAlignment`, `OfficeTextPlacement`, `OfficeTextZoneLayout`, `OfficeTextZone`, `OfficeRichTextRun`, `OfficeRichTextSegment`, `OfficeRichTextLine`, `OfficeRichTextBlockLayout`, `OfficeTextBlockRenderer.DrawRasterRichTextBlock`, `OfficeTrueTypeFont.TryLoadFontFamily`, and `OfficeGeometry` for shared dependency-free word wrapping, rich-run tokenization, stacked rich-run text-element layout, long-word breaking, hard-break normalization, max-line measurement, single-line end/start trim-to-width, shrink-to-fit font sizing for measured single-line and rich-run text, bounded wrapped text fit, bounded plain/rich text-block layout orchestration, visible-height text-block clipping with ellipsis, clipped-state reporting, reusable top/anchor/line-left placement, reusable three-column text zone layout, raster rich-run placement/render dispatch, dependency-free font-family fallback resolution, family-aware raster text measurement/rendering, shared point rotation for rotated text placement, reusable point distance, angle conversion, raster-space point rotation, and polyline-by-length interpolation. `OfficeIMO.Visio` consumes those helpers for both native PNG and SVG text blocks instead of keeping private wrap/break/measure/fit/placement implementations per output format, and PNG/SVG connector label placement plus label-collision layout now use shared Drawing geometry instead of private interpolation copies, while Visio-specific enum mapping, rotation, styling, label-background behavior, underline drawing, SVG text emission, and page-coordinate mapping remain in the Visio adapter. `OfficeIMO.Excel` now uses the same shared line model and `LayoutTextBlock` coordinator for plain wrapped, hard-break, shrink-to-fit, clipped, and forced-single-line cell text in both PNG and SVG export; Excel rich text now uses `LayoutRichTextBlock` plus `DrawRasterRichTextBlock` for hard-break, wrapped, shrink-to-fit, and basic rotated run-preserving PNG output, and `LayoutStackedRichTextBlock` for basic stacked run-preserving PNG/SVG output, while still emitting rotation approximation diagnostics. Excel page-sliced header/footer text zones now use shared Drawing zone layout plus start/end ellipsis trimming while supported font-family/style/color/size/decorations flow through shared rich text runs, shared family-aware rich text measurement, shared raster rich-text rendering, and SVG segment output. Drawing-level tests cover the shared wrapping/trim/fit/placement/clipping/shrink/orchestration/rich-run/font-family-measurement/rotation-placement/text-zone/polyline-interpolation/angle-conversion/raster-rich-renderer/stacked-rich contracts, the Excel image export tests prove the public Excel surface renders through the current renderer, `VisioSvgExport` proves hard-break and bounded SVG text output, `VisioPngExport` proves native PNG text scenarios still pass, and the native premium Visio baseline suite proves the gallery visuals remain approved.

Latest Excel render-plan checkpoint: `OfficeIMO.Drawing.OfficeTextBlockRenderPlan` now supports left/top rectangle-based placement and measured plain/stacked text-block creation, not only Visio-style center-based placement. Excel plain cell text builds that shared render plan once for PNG and SVG, then passes the resolved layout rectangle, alignment, vertical placement, wrapping, shrink-to-fit sizing, and stacked-text layout into the renderer-specific emission path. Excel still owns workbook style semantics, font-family requests, rotation policy, clipping diagnostics, and rich-text policy, but the reusable text-block placement brain is now shared with Visio through Drawing.

Latest rich-text SVG checkpoint: `OfficeIMO.Drawing.OfficeTextBlockRenderer` now owns measured rich-text SVG block emission for multiline placement, run baselines, horizontal/vertical alignment, text decorations, font-family attributes, and optional rotation metadata. Excel cell rich text and page-sliced header/footer rich text now call that shared Drawing helper for SVG output instead of carrying local cursor/baseline loops beside the raster rich-text renderer. Excel still owns run extraction, style fallback, field parsing, clipping scopes, and diagnostics; Drawing owns the reusable SVG run-placement and segment emission path.

Latest data-bar geometry checkpoint: `OfficeIMO.Drawing.OfficeDataBarRenderer` now exposes a reusable resolved data-bar geometry contract in addition to its PNG/SVG helpers. Excel conditional data bars already render through the shared Drawing PNG/SVG paths, and PDF table data bars plus Visio data-graphic bars now consume the same Drawing-owned ratio/clamping geometry while keeping their native PDF stream and VSDX shape emission in their adapters. This keeps proportional bar placement in one rendering brain without forcing every target format through the same output primitive.

Latest Bezier geometry checkpoint: `OfficeIMO.Drawing.OfficeGeometry` now owns reusable quadratic and cubic Bezier curve sampling for dependency-free path flattening. Drawing raster path rendering and Visio preserved VSDX geometry now consume the same sampled-curve primitive instead of maintaining separate curve loops, while Visio still owns VSDX row parsing, relative/absolute point extraction, spline/NURBS semantics, and page-coordinate policy. Focused Drawing raster plus Visio PNG/SVG preserved-geometry tests prove the shared curve sampler keeps the existing visible output contract.

Latest shape-transform checkpoint: `OfficeIMO.Drawing.OfficeDrawingRasterRenderer` now honors `OfficeShape.Transform` for raster output instead of leaving transform-aware rendering to SVG only. Transformed rectangle, rounded rectangle, ellipse, polygon, line, and path shapes are routed through shared transformed contour/path primitives; Excel worksheet DrawingML objects now carry authored rotation from `a:xfrm` into the neutral visual snapshot and attach it to the shared `OfficeShape` before PNG/SVG export. Excel expands rotated drawing-object overlay scenes so the shared renderer is not clipped to the unrotated shape bounds, and emits `ExcelDrawingShapeTextRotationApproximation` when rotated shape text is present. Focused Drawing raster tests prove transformed shape pixels move as expected, and Excel object tests prove rotated preset shapes render through public range PNG/SVG export with diagnostics for the remaining text approximation. A manually opened PNG artifact verified the rotated preset shape is no longer clipped, while the review also confirmed that preset geometry polish itself remains a premium fidelity item rather than a solved problem.

Latest stacked-text checkpoint: `OfficeIMO.Drawing` now owns `OfficeTextLayoutEngine.LayoutStackedTextBlock` and `LayoutStackedRichTextBlock` for upright one-text-element-per-line stacked text layout. Excel `TextRotation=255` renders plain and rich text through that shared layout in PNG and SVG with `ExcelCellTextRotationApproximation` diagnostics instead of the old unsupported stacked-text diagnostic or a rich-text layout approximation. A dedicated approved stacked-text PNG/SVG baseline proves the output is readable, upright, styled per rich run, nonblank, and free of SVG rotation transforms. Premium still needs Excel-exact stacked baseline metrics, but stacked rich text is now a shared Drawing renderer path rather than an adapter-local fallback.

Latest font diagnostics checkpoint: Excel cell, rich text run, chart text, and page-sliced header/footer image export now use Drawing's `IMAGE_FONT_SUBSTITUTED` contract. Caller-supplied TrueType faces are resolved before platform fallback, and the diagnostic carries the Excel cell, chart, or header/footer source reference. Focused tests cover missing and caller-scoped fonts across plain cells, rich text, chart titles, and formatted header/footer text.

Latest header/footer default-font checkpoint: Excel page-sliced header/footer PNG/SVG export now resolves the workbook default font family once in the Excel adapter and passes that family into the shared Drawing plain and rich text paths. Plain header/footer text and formatted runs without an explicit `&"Font"` token therefore render and measure through the workbook default family fallback list instead of a hardcoded Arial-only branch, while authored header/footer font-family tokens remain explicit adapter-owned requests. Focused header/footer image export tests now assert the workbook default font family in both plain and formatted SVG output across `net472`, `net8.0`, and `net10.0`.

Latest stroke consolidation checkpoint: `OfficeIMO.Drawing` now owns reusable parallel-line geometry plus PNG/SVG parallel styled-line emission for dependency-free double-line rendering. Excel border export still owns OpenXML border-style mapping, widths, colors, and diagnostics, but double borders now call the shared Drawing raster/SVG primitives instead of carrying private parallel-offset drawing code in the Excel renderer. Drawing-level tests pin the shared geometry, SVG output, and raster pixels, while the existing Excel premium border-style test continues to prove the public range image export surface.

Latest gradient consolidation checkpoint: `OfficeIMO.Drawing` now owns `OfficeLinearGradient.FromAngle` for normalized two-stop gradient endpoint projection from an authored angle. Excel fill export still owns OpenXML gradient stop/color extraction and unsupported-gradient diagnostics, but PNG/SVG cell fills now call the shared Drawing primitive instead of carrying private angle/vector math in the Excel renderer. Drawing-level tests pin horizontal, vertical, diagonal, positive wrap, and negative-angle endpoint projection while the Excel pattern-fill image test continues to prove the public range export surface.

Latest SVG image embedding checkpoint: `OfficeIMO.Drawing.OfficeSvgImageRenderer` now owns the embeddable image-format to MIME content-type policy for SVG `<image>` output. Excel worksheet image export still owns source extraction and diagnostics, but its SVG path asks Drawing whether PNG, JPEG, GIF, or SVG bytes can be embedded instead of carrying a private content-type switch.

Latest Visio preview content-type checkpoint: `OfficeIMO.Drawing.OfficeSvgImageRenderer` now also owns SVG-embeddable content-type resolution from declared package metadata, generic MIME values, image byte signatures, and file extensions. Visio package-preview rendering still owns VSDX relationship selection, stencil fallback policy, and page/shape placement, but no longer carries private PNG/JPEG/GIF/SVG content-type sniffing or XML-preamble SVG detection beside Drawing.

Latest image MIME policy checkpoint: `OfficeIMO.Drawing.OfficeImageInfo` now owns MIME-to-`OfficeImageFormat` normalization, including parameter stripping and common JPEG/SVG/metafile aliases. Excel worksheet/header-footer image insertion, Visio package asset/content-type defaults, and Word custom image part plus inspection content-type mapping consume that shared format policy, while Excel, Visio, and Word still own OpenXML/VSDX-specific package part creation, relationship wiring, and document-surface enum adapters.

Latest SVG text-measurement checkpoint: Visio SVG text export now routes bounded text wrapping through `OfficeIMO.Drawing.OfficeTextMeasurer` instead of carrying a private character-width heuristic beside the shared text layout engine. Visio still owns Visio text-box geometry, margins, rotation, background-label behavior, and SVG attributes, while Drawing owns the deterministic dependency-free fallback metrics used to decide line breaks. A public Visio SVG export test computes expected line breaks with the shared measurer and compares them to emitted `<tspan>` lines, pinning the adapter to the shared measurement contract.

Latest Visio resize-text checkpoint: Visio `ResizeToText` and `ResizeLabelToText` now consume `OfficeTextLayoutEngine.WrapLines` for measured line construction, tab-separated word boundaries, hard breaks, and long-word breaking instead of carrying a private wrap helper beside the renderers. Visio still owns shape/connector label sizing policy, margins, padding, minimum/maximum dimensions, and inch-to-pixel conversion. Focused Drawing tests pin tab-aware shared wrapping, public Visio layout tests pin long connector-label words staying inside the maximum width, and existing Excel/Visio renderer-facing tests prove the shared text layout path still drives image/SVG output.

Latest connector-arrow geometry checkpoint: `OfficeIMO.Drawing` now owns reusable connector arrowhead segment selection and triangular arrowhead point generation. Visio PNG and SVG export still own page/raster/SVG coordinate conversion, arrow presence policy, color, and final raster fill or SVG path emission, but both renderers now consume the same Drawing geometry instead of duplicating `atan2`/wing-angle arrowhead math and collapsed-segment skipping. Drawing-level tests pin the shared arrowhead geometry and terminal-segment behavior, while existing Visio PNG/SVG connector-arrow tests continue to prove the public export surfaces.

Latest connector-endpoint geometry checkpoint: `OfficeIMO.Drawing` now owns reusable rectangle-boundary endpoint resolution for fallback connector routing. Visio PNG and SVG export still own shape/page bounds extraction, connection-point semantics, waypoints, right-angle policy, and coordinate conversion, but both renderers now consume the same Drawing boundary rule instead of carrying duplicate center-delta side-selection math. Drawing-level tests pin horizontal, vertical, unordered-bound, and aligned-center tie behavior so future Excel/Office connector rendering can reuse the same endpoint primitive.

Latest connector-polyline geometry checkpoint: `OfficeIMO.Drawing` now owns reusable connector polyline construction from endpoints, explicit waypoints, and right-angle fallback elbows. Visio PNG/SVG connector rendering, connector label layout, and save-time connector label placement consume it, while Visio still owns endpoint extraction, connection-point semantics, waypoint source policy, coordinate conversion, and document persistence behavior. Drawing-level tests pin explicit-waypoint precedence, right-angle fallback shape, straight-line fallback, and the `OfficePoint` overload so future Excel/Office connector rendering can reuse the same route-building primitive without growing another rendering brain.

Latest connector-path consumer checkpoint: Visio inspection snapshots, visual-quality analysis, connector-label overlap layout, and orthogonal route scoring now consume `OfficeIMO.Drawing` connector polyline, interpolation, and distance primitives instead of carrying private route-building or by-length interpolation copies. Visio still owns the coordinate-space decisions, endpoint extraction, shape/label collision policy, routing candidate generation, and page/document semantics; Drawing owns the generic path math shared by renderers, save-time label placement, inspection, quality, layout, and tests.

Latest segment-intersection geometry checkpoint: `OfficeIMO.Drawing` now owns reusable segment/segment and segment/rectangle intersection tests, including boundary-touch and collinear-overlap behavior. Visio render-label collision checks, connector-label overlap layout, visual-quality connector-crossing analysis, orthogonal route scoring, and route-crossing tests now call the shared geometry primitive instead of carrying separate orientation/on-segment implementations. Visio still owns shape bounds, endpoint selection, collision policy, and source diagnostics; Drawing owns the generic geometry contract future Excel/Office connector and object rendering can reuse.

Latest raster polygon adapter checkpoint: `OfficeIMO.Drawing.OfficeRasterCanvas` now accepts tuple-point polylines, polygon fills, even-odd contours, and styled polygon outlines directly. The remaining Visio PNG `RasterCanvas` adapter no longer owns tuple-to-point conversion or closed styled-polygon stroking; it delegates those reusable raster primitives to Drawing while keeping Visio-owned raster coordinate conversion, page rotation mapping, and shape/stencil policy in the adapter.

Latest image-placement checkpoint: `OfficeIMO.Drawing` now owns the shared `OfficeImagePlacement` fit rectangle primitive for `Stretch`, `Contain`, and `Cover`. Visio package-preview PNG rendering and PDF page-background/header-footer image rendering now consume the same dependency-free placement math instead of carrying separate aspect-ratio branches, while Visio still owns package-preview policy and PDF still owns PDF image streams, source-crop semantics, and clipping policy. Drawing-level tests pin the shared fit rectangles so future Excel, worksheet image, chart snapshot, PowerPoint, and Word rendering paths can reuse one image-placement brain.

Latest source-crop checkpoint: `OfficeIMO.Drawing` now owns `OfficeImageSourceCrop` for normalized image-crop edge fractions, crop presence, visible source width/height, and collapsed authored-crop clamping. Excel worksheet image snapshots expose that shared crop value while preserving their existing crop ratio accessors, and Excel PNG/SVG image rendering plus PDF source-cropped image placement now consume the shared visible-source ratios instead of carrying separate crop-width/height math. PDF still owns its stricter public `PdfImageSourceCrop` validation and bottom-left image-stream coordinate policy, while Drawing owns the generic source-crop contract future worksheet images, PowerPoint pictures, Word/VML pictures, and shared image projection helpers can reuse.

Latest image-projection checkpoint: `OfficeIMO.Drawing` now owns `OfficeImageProjection` as the shared render intent for destination placement, normalized source crop, rotation center, rotation angle, and horizontal/vertical flips. `OfficeRasterCanvas.DrawImage`, `OfficeSvgImageRenderer.AppendImage`, and `OfficeSvgImageRenderer.WriteImage` all accept the projection; Excel worksheet image PNG/SVG rendering now builds one projection per visual image instead of maintaining parallel raster and SVG argument lists for crop, transform, scale, and placement; and Visio package-preview PNG/SVG artwork now routes its fitted/rotated preview images through the same projection contract. Excel still owns worksheet anchoring, selected-range clipping policy, content-type diagnostics, and image-format support decisions, and Visio still owns page coordinate conversion, package-preview discovery, and preview policy, while Drawing owns the reusable projection contract future PDF image stamps, PowerPoint pictures, Word pictures, and chart snapshot image overlays can consume.

Latest PDF image checkpoint: PDF image extraction now uses `OfficeIMO.Drawing` for PNG container, chunk, CRC, and zlib scanline writing instead of carrying a private PNG writer in the PDF resource resolver. PDF still owns PDF stream decoding, soft-mask composition policy, color-space decisions, and image resource semantics, while Drawing owns the reusable dependency-free PNG byte contract.

Latest sparkline checkpoint: Excel authored sparklines are now discovered through a shared worksheet sparkline resolver used by feature reporting, PDF preflight, and image export. Same-sheet numeric line, column, and win/loss sparklines render in PNG/SVG through the Excel visual snapshot and shared Drawing primitives, including basic series colors, negative colors, markers, and zero-axis output. Rendered sparklines emit `ExcelSparklineRenderingApproximation`, and cross-sheet or unresolved sparkline data still emits stable source diagnostics instead of being hidden. A dedicated approved PNG/SVG sparkline visual baseline now gates line, column, and win/loss output through the shared Drawing-backed baseline comparison helper.

Latest object checkpoint: Excel comments/notes and threaded comments now produce visible top-right cell indicators in PNG/SVG when their target cell is inside the exported range. When `ShowCommentBodies` is enabled, visible classic and threaded comment bodies also enter the neutral Excel snapshot and render as dependency-free callouts with anchored pointers through shared Drawing shapes, shared text layout, and shared Drawing text-block emission for callout body text; enabled bodies emit `ExcelCellCommentBodyApproximation` or `ExcelThreadedCommentBodyApproximation` instead of unsupported-body diagnostics. With the option disabled, `ExcelCellCommentUnsupported` and `ExcelThreadedCommentUnsupported` remain stable source diagnostics so compact indicator-only exports stay explicit. Simple worksheet rectangle/rounded-rectangle drawing shapes with solid RGB fill/outline and plain text now enter the neutral Excel snapshot and render in PNG/SVG through `OfficeIMO.Drawing`; supported shapes, images, opt-in comment bodies, and charts now share explicit visual layers instead of growing separate renderer brains. Unsupported geometry, theme/system/transformed fills, rotated shapes, connectors, group shapes, non-chart graphic frames, and Excel-exact comment popover geometry/state remain diagnosed follow-up work. The premium Excel baseline now includes the legacy red comment indicator, a dedicated drawing-object approved PNG/SVG baseline gates the first shape/text-box slice, focused object tests prove shape-over-image and image-over-shape output using source order, SVG order, and decoded PNG pixels, and comment-body tests prove drawing-layer placement, decoded callout pixels, and SVG text/color/pointer output.

Latest image checkpoint: Excel range export now includes worksheet images whose visual rectangle intersects the selected range even when the image anchor cell is outside that range. Raster output relies on the shared Drawing canvas bounds for clipping; SVG output emits explicit range clip paths for embedded images. Two-cell anchored worksheet pictures now derive visual width and height from their OpenXML from/to marker geometry instead of falling back to the embedded image's natural pixel size. Authored picture crop rectangles (`a:srcRect`), basic picture rotation, and horizontal/vertical flips now flow from Excel image metadata into the neutral visual snapshot and render in PNG/SVG through one shared Drawing image projection path. Focused contract tests prove decoded PNG pixels, SVG clip/transform structure, two-cell marker sizing, cropped picture output, visible rotated image output, and combined crop-plus-flip-plus-rotation output, and dedicated approved clipped-image, two-cell image, cropped-image, rotated-image, and transformed-image PNG/SVG baselines make the behavior visually reviewable.

Latest chart checkpoint: Excel chart snapshots now carry a first slice of authored chart layout/style into the shared `OfficeIMO.Drawing` chart renderer instead of flattening everything to defaults. Chart area solid fill/outline/width/preset dash, plot area solid fill/outline/width/preset dash, simple authored series fill/line colors, line widths, and preset dashes, simple point fills, simple marker fill colors, marker visibility, simple marker size, simple marker solid outline color/width, simple circle/square/diamond/triangle/dash/dot/plus/X/star marker shapes, simple category/value major-gridline color/visibility/width/preset dash, simple category/value minor-gridline color/visibility/width/preset dash, simple category/value axis-line color/visibility/width/preset dash, simple category/value major tick marks, simple category/value minor tick marks, category/value axis label visibility when Excel tick-label position is `none`, simple high/low/next-to category/value tick-label placement, simple maximum-crossing horizontal category-axis and vertical value-axis placement, simple category/date-axis reverse-order rendering, simple value-axis number formats for vertical and horizontal bar orientations, simple value-axis display-unit scaling and labels, simple linear value-axis min/max/major/minor-unit scaling, simple title text color/font-family/font-size/bold/italic, simple legend text color, simple data-label text color, simple axis-label text color, simple axis-title text-color override, simple legend/data-label/axis-label font sizes, simple axis-title font size, simple legend/data-label/axis-label font families, simple legend/data-label/axis-label bold/italic buckets, simple axis-title font-family/bold/italic overrides, legend presence/position/overlay, title overlay, category/value axis titles, and chart-level data-label flags/position/number format now flow into shared `OfficeChartStyle`, `OfficeChartSeries`, and `OfficeChartLayout`. Unsupported or approximate chart details still emit stable source diagnostics, including `ExcelChartTrendlineUnsupported`, `ExcelChartDataLabelPointOverridesApproximated`, `ExcelChartDataLabelLeaderLinesUnsupported`, `ExcelChartAreaStyleApproximation`, `ExcelChartGridlineStyleApproximation` for complex gridline effects, `ExcelChartAxisStyleApproximation`, `ExcelChartAxisTickLabelPositionApproximation`, `ExcelChartAxisMinorTickMarkPlacementApproximation` for remaining approximate minor tick-mark placement, `ExcelChartAxisCrossingApproximation`, `ExcelChartAxisScaleApproximation`, `ExcelChartAxisNumberFormatApproximation`, `ExcelChartCategoryAxisNumberFormatUnsupported`, `ExcelChartTextStyleApproximation`, and `ExcelChartSeriesStyleApproximation` when chart styling goes beyond the supported simple area/series/point/marker/gridline/axis/chart-text color/font-family/font-size/font-style/line-width/preset-dash/axis-number-format/display-unit/label-visibility/major/minor-tick-mark slice or when Excel asks for conflicting text colors, font families, font sizes, or font styles inside one supported text bucket. Focused contract tests prove the snapshot bridge, SVG style/text output, high-scale PNG visual pixels, authored chart/plot area border width and preset dash output, authored series-color, series-line-width, and series-line-dash SVG/PNG output, authored point-color SVG/PNG output, simple marker-fill/size/shape/outline including line-based dash/dot/X/star markers, simple category/value major-gridline color/width/dash SVG/PNG output, simple value-axis minor-gridline color/width SVG/PNG output, suppressed gridline output, simple category/value axis-line color/width/dash SVG/PNG output, suppressed axis-line output, suppressed category/value axis labels, simple title/legend/data-label/axis-label/axis-title text-color SVG/PNG output, simple title font-family/font-size/italic SVG output, simple legend/data-label/axis-label/axis-title font-size SVG output, simple legend/data-label/axis-label font-family/style and axis-title font-family/style SVG output, vertical and horizontal value-axis number-format SVG output, simple display-unit label SVG output, simple linear value-axis min/max/major/minor-unit SVG output, simple major and minor axis tick-mark rendering, and trendline/chart-area/gridline/axis-placement/axis-scale/axis-number-format/category-axis-number-format/text-style diagnostics. Premium chart export is still not done: picture markers and richer marker outline effects, custom/richer series dash and effect styling, richer point-level overrides, trendlines, leader lines, point-level label overrides, remaining Excel-exact minor-gridline/tick placement edge cases, custom dash/effect parity beyond preset gridline and axis lines, chart/plot area effects beyond simple solid RGB fill/outline/width/preset dash, remaining tick-label placement edge cases beyond simple high/low/next-to/none, axis-crossing geometry beyond simple horizontal category-axis and vertical value-axis maximum crossing, log/value-axis-reverse-order/non-value-axis-unit/non-default cross-between axis geometry, Excel-exact display-unit placement/typography, richer chart title typography/effects beyond simple font-family/font-size/bold/italic, per-element chart rich text runs beyond the supported shared buckets, full custom/date/scientific/conditional tick formatting including category/date-axis tick formats, and Excel-exact chart geometry remain explicit follow-up work.

Latest SVG primitive checkpoint: `OfficeIMO.Drawing` now owns `OfficeSvgPrimitiveWriter` for dependency-free `XmlWriter` circle, rectangle, line, and path emission with shared number formatting, fill/stroke color output, rounded caps, and rounded joins. Visio built-in stencil metadata artwork now supplies only stencil geometry, placement, opacity, and shape-coordinate policy while consuming the shared Drawing writer for generic primitive output; Visio shape/background/text/connector paths also use shared `OfficeSvgFormatting.WriteColorAttribute` for writer color attributes. Focused Drawing writer tests and existing Visio stencil SVG tests prove the shared primitive output and rotated stencil artwork still work.

Latest Visio preview SVG escaping checkpoint: Visio stencil preview gallery HTML and generated thumbnail SVG now route text/attribute escaping through `OfficeSvgFormatting.Escape` instead of carrying private HTML/XML escape helpers in the Visio adapter. Visio still owns package preview extraction, gallery structure, file naming, and thumbnail layout, while Drawing owns dependency-free XML/SVG escaping for renderer-facing string builders. The browser-renderable stencil thumbnail artifact test now proves angle brackets, ampersands, and quotes are escaped in both the HTML index and generated thumbnail SVG.

Latest nested SVG consolidation checkpoint: `OfficeSvgFormatting` now owns SVG-root inner-content extraction and nested SVG viewport emission for dependency-free `StringBuilder` adapters. Excel chart SVG overlays, supported drawing-object SVG overlays, and opt-in comment body SVG callouts now use the shared helper instead of hand-assembling child `<svg>` wrappers and `viewBox` attributes in separate renderer partials. Focused Drawing formatter tests plus Excel chart, drawing-object, comment-body, drawing-object baseline, and premium range baseline tests prove the migration is behavior-preserving.

Latest SVG polygon consolidation checkpoint: `OfficeSvgFormatting` now owns complete SVG polygon element emission for `StringBuilder` renderers, including shared point-list formatting and optional fill/stroke attributes. `OfficeDrawingSvgExporter` polygon shapes and Excel comment indicator/comment-body pointer SVG output now use that shared helper while keeping Drawing shape placement and Excel comment geometry/color policy in their adapters. Focused Drawing formatter/exporter tests plus Excel comment indicator, comment body, threaded-comment indicator, and premium range baseline tests prove the migration is behavior-preserving.

Latest SVG line consolidation checkpoint: `OfficeSvgFormatting` now owns complete SVG line element emission for `StringBuilder` renderers, including shared coordinate formatting, stroke paint, opacity, width, dash array, and line-cap attributes. `OfficeDrawingSvgExporter` line shapes and Excel range border SVG lines now use that shared helper while keeping Drawing transform policy and Excel border-style policy in their adapters. Focused Drawing formatter/exporter tests plus Excel border-style and premium range baseline tests prove the migration is behavior-preserving.

Latest SVG rectangle consolidation checkpoint: `OfficeSvgFormatting` now owns complete SVG rectangle and rounded-rectangle element emission for `StringBuilder` renderers, including shared coordinate, size, and corner-radius formatting with adapter-supplied paint/transform attributes. `OfficeDrawingSvgExporter` rectangle shapes, Excel range gridline and cell-fill SVG rectangles, shared data-bar SVG rectangles, and shared sparkline column/win-loss SVG rectangles now use the shared helper while keeping Drawing transform policy and Excel style/conditional-formatting policy in their adapters. Focused Drawing formatter/exporter/data-bar/sparkline tests plus Excel border-style, pattern-fill, conditional-formatting, sparkline, and premium range baseline tests prove the migration is behavior-preserving.

Latest SVG sparkline primitive consolidation checkpoint: `OfficeSvgFormatting` now owns complete SVG polyline and circle element emission for `StringBuilder` renderers, including shared point-list, center/radius, and fill formatting. `OfficeSparklineRenderer` line-series SVG paths and marker SVG circles now use the shared helpers while preserving the approved sparkline SVG attribute order, sparkline scaling, marker color, and series policy in the shared sparkline renderer. Focused Drawing formatter/sparkline tests plus Excel sparkline and premium range baseline tests prove the migration is behavior-preserving.

Latest SVG ellipse consolidation checkpoint: `OfficeSvgFormatting` now owns complete SVG ellipse element emission for `StringBuilder` renderers, including shared center/radius and fill opacity formatting. `OfficeDrawingSvgExporter` ellipse shapes now use the shared helper while keeping Drawing-owned placement, paint, and transform policy in the exporter. Focused Drawing formatter/exporter tests prove the migration is behavior-preserving.

Latest SVG path-element consolidation checkpoint: `OfficeSvgFormatting` now owns complete SVG path element emission for `StringBuilder` renderers, including shared path data serialization, `d` attribute escaping, and adapter-supplied paint/transform attributes. `OfficeDrawingSvgExporter` freeform path shapes and path clip definitions now use the shared helper while keeping Drawing-owned placement, clip, paint, and transform policy in the exporter. Focused Drawing formatter/exporter tests prove the migration is behavior-preserving.

Latest SVG clip-rectangle consolidation checkpoint: `OfficeDrawingSvgExporter` clip-path rectangle and rounded-rectangle definitions now use `OfficeSvgFormatting.AppendRectElement` instead of local element assembly, keeping clip ownership in Drawing while sharing coordinate, size, and corner-radius serialization with every other `StringBuilder` SVG rectangle consumer. Focused Drawing clip-path/exporter tests prove the migration is behavior-preserving.

Latest SVG positioned-text consolidation checkpoint: `OfficeTextBlockRenderer` now owns positioned SVG text/tspan element emission for callers that already resolved anchor coordinates, first baseline, line height, font, and style. `OfficeDrawingSvgExporter` text boxes now use the shared helper while keeping Drawing-owned text-box placement and font/style policy in the exporter, and text fill opacity now uses the same shared SVG paint formatting as other Drawing primitives. Focused Drawing renderer/exporter tests prove the migration is behavior-preserving apart from the intentional alpha-opacity improvement.

Latest Excel rotated SVG text consolidation checkpoint: `OfficeTextBlockRenderer.AppendSvgTextElement` now supports positioned underline and rotation attributes, so Excel's plain rotated cell-text SVG path uses the shared Drawing text helper instead of private `<text>` assembly. Excel still owns rotation interpretation, clipping, alignment, font/style/color resolution, and approximation diagnostics; Drawing owns the text element, text-anchor, fill opacity, style attributes, escaping, and rotate-transform emission. Focused Drawing helper tests plus the public Excel rotated PNG/SVG text test prove the migration is behavior-preserving.

Latest Excel rich-text SVG segment consolidation checkpoint: `OfficeTextBlockRenderer` now owns SVG text element emission for measured rich text segments through `AppendSvgRichTextSegment`. Excel rich cell text still owns run extraction, style fallback, line layout, cursor placement, clipping, rotation grouping, and diagnostics, but each rendered SVG segment now uses the shared Drawing helper for element assembly, escaping, text-anchor, fill opacity, font, and bold/italic/underline attributes. Focused Drawing helper tests plus the public Excel rich-text SVG test prove the migration is behavior-preserving.

Latest Excel comment-title SVG consolidation checkpoint: Excel comment-body title text now uses `OfficeTextBlockRenderer.AppendSvgTextElement` instead of private `<text>` assembly. Excel still owns comment-body placement, callout geometry, clipping, colors, source references, and approximation diagnostics, while Drawing owns title text element assembly, escaping, text-anchor, fill, font, and bold style emission. The public comment-body SVG/PNG object-export contract now proves the title remains visible and styled through the shared helper.

Latest Visio stencil-thumbnail SVG text consolidation checkpoint: Visio stencil preview thumbnails now use `OfficeTextBlockRenderer.AppendSvgTextElement` for their caption text instead of private SVG text assembly. Visio still owns gallery discovery, thumbnail sizing, data URI embedding, and review HTML output, while Drawing owns caption element assembly, escaping, text-anchor, fill, and font attributes. The existing browser-renderable thumbnail artifact test now proves the generated caption shape.

Latest Excel SVG root-background consolidation checkpoint: Excel range SVG export now uses `OfficeSvgFormatting.AppendRectElement` for the document background rectangle instead of private root `<rect>` assembly. Excel still owns canvas dimensions, viewBox policy, and selected background color; Drawing owns rectangle geometry, number formatting, and fill/opacity emission. The basic public Excel PNG/SVG export contract now asserts the root background artifact shape.

Latest Visio stencil-thumbnail wrapper SVG consolidation checkpoint: Visio stencil preview thumbnails now use shared Drawing helpers for their wrapper rectangles and preview image element as well as caption text. `OfficeSvgImageRenderer.AppendImage` now supports adapter-supplied `preserveAspectRatio`, so Visio keeps thumbnail layout and package/gallery policy while Drawing owns rectangle geometry, image geometry, data URI construction, preserve-aspect attribute emission, and text element formatting. The browser-renderable thumbnail artifact test now proves the full wrapper shape.

Latest Excel top/bottom conditional-formatting checkpoint: Excel image export now renders bounded numeric top/bottom count and percent rules with solid differential fills, including tied values at the cutoff, through the existing conditional fill pipeline. Rule snapshots expose top/bottom rank, bottom, and percent metadata; public top/bottom builder overloads can attach a fill color. Invalid or nonnumeric top/bottom rules remain source-diagnosed with `ExcelConditionalTopBottomUnsupported` instead of being silently ignored.

Latest Excel average conditional-formatting checkpoint: Excel image export now renders above-average and below-average rules with solid differential fills through the same conditional fill pipeline, including equal-to-average inclusion when requested. Rule snapshots carry above/below/equal/std-dev metadata, and the public sheet and fluent range APIs can attach average-rule fill colors. Standard-deviation average rules and unsupported conditional-formatting families such as date/time still remain source-diagnosed instead of being silently ignored.

Latest Excel text conditional-formatting checkpoint: Excel image export now renders contains-text, not-contains-text, begins-with, and ends-with conditional-formatting rules with solid differential fills through the same conditional fill pipeline. Rule snapshots carry the rule text payload, public sheet/fluent APIs can author the text rules with fills, and malformed text rules emit `ExcelConditionalTextRuleUnsupported` instead of silently disappearing.

Latest shared chart typography checkpoint: `OfficeIMO.Drawing` now owns less-squinty default chart legend, axis, and data-label font sizes plus label boxes sized from the active font instead of fixed 10/11-pixel bands. Excel chart export remains a thin adapter that maps workbook-authored chart text styles into the shared chart snapshot, while default chart text readability improves for Excel PNG/SVG visual output and future Drawing consumers without adding an Excel-only chart rendering branch. The Excel image visual baselines were regenerated through `OFFICEIMO_UPDATE_EXCEL_IMAGE_BASELINES=1` after visual inspection, and the focused baseline-producing tests now pass against the refreshed approved PNG/SVG artifacts.

Latest PDF extracted-image MIME checkpoint: PDF image extraction now uses `OfficeImageInfo.GetMimeType` from `OfficeIMO.Drawing` for reconstructed JPEG/PNG file identities instead of hardcoded MIME strings in the PDF reader. PDF still owns PDF stream/filter interpretation, soft-mask reconstruction, and PNG file construction, while Drawing remains the central image-format vocabulary shared by Excel, Word, Visio, SVG embedding, and PDF extraction.

Latest WebP image-format checkpoint: `OfficeIMO.Drawing` now recognizes WebP as a shared image format for MIME, extension, and SVG data-URI embedding policy. Visio stencil preview gallery image fallback now delegates extension-to-MIME resolution to Drawing instead of keeping a private switch, so package-backed preview artwork and future image export paths share the same image vocabulary.

Latest PDF image content-type checkpoint: `OfficeImagePdfCompatibility` now owns the first-party PDF image MIME support contract for PNG/JPEG declarations. Excel PDF export and Excel PDF preflight both consume that shared Drawing policy instead of carrying parallel `image/png`/`image/jpeg` checks, while Excel keeps its own warning wording and source-specific diagnostics.

Latest PowerPoint image-extension checkpoint: PowerPoint path-based picture, background, and poster-image entrypoints now resolve image file extensions through `OfficeImageReader.FromExtension` and a small PowerPoint-owned adapter to `ImagePartType`. PowerPoint still owns OpenXML image part creation, relationship wiring, and media placement, while Drawing owns the shared extension-to-image-format vocabulary.

Latest shared SVG-embeddable MIME checkpoint: `OfficeSvgImageRenderer` now normalizes declared MIME content types through the same SVG-embeddable image policy it already applies to detected image formats. Visio package-preview artwork uses that shared policy for browser-renderable preview decisions instead of carrying a private PNG/JPEG/GIF gate, so sniffed SVG previews with XML preambles render through the shared SVG image path. PowerPoint's default package thumbnail MIME value also comes from `OfficeImageInfo`, keeping built-in JPEG identity on the shared Drawing vocabulary while PowerPoint remains responsible for OpenXML thumbnail part creation.

Latest Excel default image-MIME checkpoint: Excel's public image APIs keep their friendly compile-time `"image/png"` optional defaults, but internal template, URL-image, and header/footer fallback defaults now call `OfficeImageInfo.GetMimeType(OfficeImageFormat.Png)`. Excel remains responsible for workbook insertion and OpenXML part wiring while Drawing owns the canonical PNG MIME value used across document families.

Latest Excel page-setup image checkpoint: page-sliced worksheet PNG/SVG export now composes rendered worksheet-page content onto a physical page canvas through shared `OfficeImageComposer` layers. The Excel adapter owns OpenXML page semantics and now applies orientation, margins, manual scale, supported worksheet paper-size codes, and bounded one-page fit-to-width/fit-to-height scaling for manual-page-sliced multi-output image exports while `OfficeIMO.Drawing` owns neutral physical page sizes and the actual PNG/SVG page composition. Missing paper size emits `ExcelPageSetupPaperSizeDefaulted`; unknown paper-size codes emit `ExcelPageSetupPaperSizeUnsupported` and fall back to Letter. Fit-to-width/fit-to-height values above one page in either dimension remain explicitly diagnosed with `ExcelPageSetupUnsupported` until automatic multi-page fit pagination is implemented instead of being faked.

Latest page-layout visual checkpoint: page-sliced worksheet PNG/SVG export now has a dedicated approved visual baseline for a physical Letter landscape page. The scenario exports through the public worksheet image path, proves manual page-break slicing, repeated print-title rows, supported fit-to-width page setup, rendered header/footer text chrome, stable formatting-approximation diagnostics, and decoded nonblank page content on the shared Drawing canvas. Automatic multi-page fit pagination, large-sheet tiling, broader paper-size coverage, and Excel-exact header/footer image behavior remain premium work.

Latest header/footer image checkpoint: page-sliced worksheet export renders `&G` header/footer images through the shared Drawing composition and decode policy. SVG output embeds direct-safe formats or transcodes shared/caller-decoded rasters, while raster output uses the shared decoder and `ImageCodec` boundary. Undecodable sources render a visible placeholder with `IMAGE_SOURCE_DECODE_FALLBACK` instead of silently disappearing or emitting an Excel-only skip code. The layout remains an approximation with stable `ExcelHeaderFooterImageApproximation` diagnostics; Excel-exact header/footer image placement and scaling-with-document semantics remain premium work.

Latest Excel/PDF page-geometry consolidation checkpoint: `OfficeIMO.Drawing.OfficePageSize` now exposes point conversion in addition to pixel conversion, and `OfficeIMO.Excel` owns `ExcelPageSetupGeometry` as the neutral worksheet page-size, fit-scale, and margin helper. Page-sliced Excel PNG/SVG output and first-party Excel PDF output now consume the same worksheet paper-size resolver instead of maintaining separate OpenXML paper-size maps. Excel PDF keeps explicit `ExcelPdfSaveOptions.PageSize` precedence, but when no explicit PDF page size is supplied it now honors supported worksheet paper-size codes such as A4 through the shared resolver.

Latest PDF image-placement consolidation checkpoint: `OfficeIMO.Drawing.OfficeImageRenderPlan` now owns shared target-box, fit, and source-crop placement math for image rendering, with top-left and bottom-left coordinate-system entrypoints. `OfficeIMO.Pdf` consumes that plan for flow, table, header/footer, canvas, and page-background image placement instead of carrying private crop-plus-fit or page-fit algorithms in the PDF writer. PDF still owns PDF clip paths, image XObject streams, tagging, annotations, page ordering, opacity graphics states, and compression, but the reusable visual placement math now lives in Drawing for Excel, Visio, PDF, and later adapters to share. PDF link-annotation bounds for cover-fitted, clipped, or source-cropped images are now centralized in one PDF-owned helper shared by flow and table-cell image rendering.

Latest DrawingML preset-geometry consolidation checkpoint: `OfficeIMO.Drawing.OfficeShapePresets` now owns the richer dependency-free DrawingML preset geometry for heart, cloud, donut, can, cube, and left-right arrow in addition to the simpler existing presets. Word native PDF rendering now passes the serialized OpenXML preset token into the shared Drawing preset table instead of carrying a private `CreateNativeDrawingPresetShape` geometry switch and local path helpers. PowerPoint PDF and future Excel object rendering therefore share the same preset vocabulary, while Word/PDF remains responsible only for WordprocessingML extraction, dimensions, and style application. Focused Drawing and Word/PDF tests prove the shared preset geometry, the line versus straight-connector contract, and DrawingML preset PDF rendering on `net8.0` and `net472`.

Latest Excel worksheet DrawingML preset checkpoint: Excel worksheet drawing-object export now validates authored `a:prstGeom` values through `OfficeShapePresets`, carries the serialized preset token and flip flags in the neutral visual snapshot, and renders supported presets through shared Drawing geometry. Excel no longer owns a local rectangle/rounded-rectangle geometry switch for supported worksheet shapes; it still owns worksheet anchors, fill/outline/text extraction, and diagnostics for unsupported colors, rotation, groups, connectors, and missing or unsupported geometry. Focused object tests prove a `heart` preset exports as SVG path geometry and decoded PNG fill pixels through the public range export path without `ExcelDrawingShapeUnsupported`.

Latest rotated preset visual checkpoint: the shared `OfficeShapePresets` heart geometry was polished in `OfficeIMO.Drawing` and is now covered by a dedicated Excel approved PNG/SVG baseline for a rotated DrawingML preset shape. The baseline exports through the public Excel range image path, asserts the SVG transform/path/fill/outline artifacts, rejects unsupported-shape and rotated-text diagnostics for the empty-text shape case, decodes the approved PNG for nonblank fill pixels, and was manually reviewed after moving the fixture layout so the object no longer collides with title or caption text. This proves the current Excel object path is using the central Drawing preset/transform renderer, while richer Excel-exact preset geometry, theme/transformed colors, connectors, grouped objects, and rotated shape text metrics remain premium work.

Latest vertical shape-text visual checkpoint: simple DrawingML vertical shape text now has a dedicated approved Excel PNG/SVG baseline. The scenario exports through the public range image path, rejects unsupported-shape and unsupported-vertical-text diagnostics, asserts that SVG output emits separate stacked letter text nodes instead of one horizontal word, decodes the approved PNG for visible shape fill and dark text pixels, and was manually reviewed as a readable stacked label. Complex vertical variants still remain diagnosed until their semantics are implemented deliberately.

## Goal

Build a dependency-free OfficeIMO image conversion stack that can render selected Office content to PNG and SVG in a deterministic, server-safe way.

The first delivered slice should be small: Excel range to PNG/SVG. The architecture must be large enough to grow into worksheet, workbook, drawing, chart, PowerPoint slide, Word page, and other Office visual exports without rewriting the foundation.

## North Star

OfficeIMO should be able to answer this family of requests:

```csharp
sheet.Range("A1:D12").SaveAsPng("range.png");
sheet.Range("A1:D12").SaveAsSvg("range.svg");

sheet.SaveAsPng("worksheet.png");
workbook.SaveAsImages("output-folder");

drawing.SaveAsPng("drawing.png");
presentation.Slides[0].SaveAsPng("slide.png");
```

The API can grow by document family, but the rendering foundation should stay shared:

1. Document package readers create a neutral visual snapshot.
2. Shared drawing/rendering code paints that snapshot.
3. Encoders write PNG or SVG.
4. Diagnostics explain every unsupported or approximated feature.

## Guardrails

- No runtime dependency on Excel, LibreOffice, browsers, Playwright, Poppler, ImageSharp, SkiaSharp, System.Drawing.Common, or platform graphics APIs.
- No product path that renders to PDF and then rasterizes the PDF with an external executable.
- No Excel-only one-off renderer hidden in `OfficeIMO.Excel`.
- No document-specific private raster canvas, PNG encoder, or PNG decoder when equivalent shared `OfficeIMO.Drawing` capability exists or can reasonably be promoted there.
- No new premium rendering work that deepens a second renderer brain before the shared Drawing migration path is explicit.
- No promise of byte-identical screenshots from desktop Office. The contract is deterministic OfficeIMO rendering with clear diagnostics and professional visual fidelity.
- No feature silently disappears. Unsupported visuals must be reported through diagnostics.
- No public API shape that blocks multi-page, multi-sheet, or multi-format exports later.

## Renderer Consolidation Findings

Current product rendering ownership should be treated as:

- `OfficeIMO.Drawing`
  - Target shared engine for dependency-free raster buffers, raster canvas operations, shared styled strokes, PNG read/write, shared chart/drawing rendering, and image export diagnostics.
  - Current branch has the new Excel-facing raster stack and first Visio-needed primitives.
- `OfficeIMO.Excel`
  - Thin document adapter over Excel visual snapshots and `OfficeIMO.Drawing`.
  - Should not gain a private pixel engine, and should keep moving reusable text/object/layout primitives into Drawing or shared Excel utilities instead of growing image-only helper brains.
- `OfficeIMO.Visio`
  - Has the strongest current native PNG renderer and visual-baseline discipline.
  - Still owns a private `RasterCanvas` for Visio-specific geometry, connector, stencil, and label-layout glue. Its private `PngRaster`, PNG encoding code, supersampled pixel buffer, alpha blending, downsample resolve, polygon/contour fill loops, line/dashed/polyline stroke loops, dashed ellipse stroke approximation, solid elliptical arcs, rotated ellipse fill/stroke, rotated/scaled image projection loop, anchored text-line drawing, fallback glyph drawing, text measurement, and PNG/SVG text wrapping helpers have been moved to shared Drawing.
- `OfficeIMO.Pdf`
  - Has PDF-specific image parsing and stream compression; this is not automatically a duplicate PNG renderer because PDF output has different stream/filter contracts.
  - Test-only PNG generation/comparison helpers now reuse shared Drawing test support.

Consolidation rule:

```text
Document package renderer = layout and source semantics only
OfficeIMO.Drawing = pixels, paths, fills, text rasterization, image decode/encode, shared diagnostics
Tests = visual comparison helpers, preferably shared once stable
```

## Current Rendering Path Inventory

This is the working map for keeping OfficeIMO on one central rendering brain.

| Path | Current owner | Shared rendering route | Adapter-owned policy | Next consolidation work |
| --- | --- | --- | --- | --- |
| Excel range/worksheet/workbook to PNG/SVG | `OfficeIMO.Excel` | Excel visual snapshots render through `OfficeIMO.Drawing` raster, SVG, PNG, text, image, chart, sparkline, data-bar, hatch, and primitive helpers. | Workbook/package extraction, range/page selection, OpenXML style/theme/number-format interpretation, cell/row/column/worksheet semantics, diagnostics source references. | Improve premium text/layout parity, richer conditional-format/style/image/chart support, page slicing, and more baseline matrices without adding Excel-only renderer primitives. |
| Visio page/package previews to PNG/SVG | `OfficeIMO.Visio` | Visio PNG/SVG text, line/stroke, raster storage, image projection, SVG primitive, nested SVG, and visual-baseline comparison paths now route through `OfficeIMO.Drawing` helpers. | VSDX page shape/connector/stencil semantics, routing, coordinate conversion, label backgrounds, gallery/package policy, optional desktop Visio validation. | Continue shrinking remaining `VisioPngRenderer.RasterCanvas` geometry glue until only Visio coordinate semantics remain. |
| PDF creation and visual PDF output | `OfficeIMO.Pdf` | Uses `OfficeIMO.Drawing` for colors, vector descriptors, charts, image helpers, shared image render-plan placement, PNG-backed visual QA, and drawing interop. PDF writer streams, filters, objects, layout, tagging, and compliance remain PDF-owned. | PDF object model, page writer, compression/filter contracts, PDF/A/PDF/UA metadata, form/signature/security/readback semantics. | Move only generic drawing math or visual QA helpers to Drawing; keep PDF stream/page writer behavior in PDF. |
| Word to PDF | `OfficeIMO.Word.Pdf` over `OfficeIMO.Pdf` | Thin adapter maps Word document semantics into PDF primitives; shared drawing behavior flows through `OfficeIMO.Pdf` and `OfficeIMO.Drawing`. | Word sections, paragraphs, lists, tables, fields, headers/footers, anchored/floating layout policy, warnings. | When Word-specific native shape/VML/chart rendering needs reusable pixels/text/path/image logic, promote that generic slice to Drawing. |
| Excel to PDF | `OfficeIMO.Excel.Pdf` over `OfficeIMO.Excel` and `OfficeIMO.Pdf` | Thin adapter reuses Excel visual/style/snapshot work, shared Excel page setup geometry, and PDF/Drawing primitives rather than owning a second spreadsheet renderer. | PDF pagination/table mapping, print-area/page setup policy, warnings, PDF-specific output shape. | Continue reducing duplication between Excel image snapshots and Excel PDF internal rendering where the shared snapshot lowers risk. |
| PowerPoint to PDF | `OfficeIMO.PowerPoint.Pdf` over `OfficeIMO.Pdf` and `OfficeIMO.Drawing` | Slide charts, simple shapes, groups, tables, and visual baselines already have direct Drawing/PDF interop. | Slide/master/layout/theme semantics, placeholder/media policy, page-sized slide mapping, warnings. | Promote generic grouped transforms, shape effects, and image projection gaps to Drawing as they become cross-document needs. |
| Visual QA | `OfficeIMO.Tests` | Excel, Visio, and PDF raster comparisons share `VisualBaselineTestSupport` backed by `OfficeIMO.Drawing` PNG read/write and diff helpers. | Scenario construction, approved artifact naming, optional external desktop/PDF tools, manual review gates. | Publish a single scenario manifest/gate for cross-document visual review and require it for premium rendering changes. |

Central ownership is now enforceable in tests:

- Rendering-capable adapters must route through `OfficeIMO.Drawing` directly or through the first-party `OfficeIMO.Pdf` engine.
- Retired Visio private PNG raster/encoding files must not be restored.
- Dependency-free rendering projects must not add ImageSharp, SixLabors.Fonts, SkiaSharp, System.Drawing.Common, browser, Office automation, or PDF-rasterization dependencies to product rendering paths.
- Excel, Visio, PowerPoint, and PDF conversion adapters must not declare product-local raster/PNG infrastructure types such as private PNG writers, PNG encoders, RGBA images, RGBA canvases, or raster render targets. Those names belong in `OfficeIMO.Drawing`; adapters can keep narrow coordinate/layout glue only when it encodes document semantics.

The current exception is intentional: `OfficeIMO.Pdf` owns PDF stream/page/writer behavior because PDF is not an image canvas. That exception does not allow PDF, Word, Excel, PowerPoint, or Visio adapters to grow private raster, SVG primitive, text layout, or PNG brains when the behavior is generic.

The first migration target is Visio PNG internals because Visio already solved several problems Excel needs for premium output: dashed strokes, rotated images, rotated/anchored text, even-odd contour fills, stencil artwork projection, and visual baseline gates. The first consolidation slices moved PNG read/write edges, supersampled pixel storage/blending/resolve, polygon fills, contour fills, line/polyline/styled stroke drawing, dashed ellipse stroke approximation, solid elliptical arcs, rotated ellipse fill/stroke, rotated/scaled image drawing, anchored text-line drawing, fallback glyph drawing, text measurement, shared PNG/SVG text wrapping/trim helpers, and rectangular raster clipping to `OfficeIMO.Drawing`; the shared visual-baseline helper now covers Excel, Visio, and PDF raster comparisons. The next slice should migrate richer path helpers while leaving document-specific Visio layout semantics in `OfficeIMO.Visio`.

## Architecture

### 1. Shared Rendering Foundation

Owner: `OfficeIMO.Drawing`

Add the missing raster side beside the existing SVG side:

- `OfficeRasterImage` or equivalent RGBA buffer.
- `OfficeRasterCanvas` for lines, rectangles, fills, clipping, and text.
- `OfficePngWriter` for dependency-free PNG encoding.
- `OfficeDrawingRasterRenderer` that renders existing `OfficeDrawing` primitives to PNG.
- Shared image export result and diagnostics types where they are not document-specific.

This is where ChartForgeX is most useful as reference material. Borrow ideas and carefully adapted code patterns from its raster image, canvas, and PNG writer. Do not depend on the ChartForgeX package.

### 2. Neutral Visual Snapshots

Owner: source document package, reusable across adapters.

Each Office domain should translate package content into a format-independent visual snapshot before writing PNG or SVG:

- Excel range snapshot.
- Excel worksheet snapshot.
- Drawing/chart snapshot.
- PowerPoint slide snapshot.
- Word page snapshot.

Snapshots should contain visual facts, not encoder decisions: bounds, cells, text runs, styles, fills, borders, images, charts, layout measurements, merges, hidden rows/columns, and source feature diagnostics.

### 3. Thin Document Adapters

Owner: document-specific packages such as `OfficeIMO.Excel`.

Document adapters should expose friendly APIs and call the shared snapshot/rendering pipeline:

- Select content.
- Build visual snapshot.
- Choose PNG or SVG renderer.
- Return bytes, stream output, or file output.
- Return diagnostics.

Adapters should not own a separate raster engine.

### 4. Diagnostics As Contract

Every export should be able to return structured diagnostics:

- Unsupported source features.
- Approximate rendering decisions.
- Missing fonts or fallback text measurement.
- Cropped content.
- Hidden rows/columns skipped.
- Images or chart snapshots that could not be rendered.

This lets the first slices be useful without pretending to be complete.

### 5. Visual Fidelity As A Product Contract

The image export goal is not merely "valid PNG/SVG bytes." Output should look like a credible Office-rendered artifact, not like a hand-built debug canvas.

Every visually exposed phase should have an explicit fidelity bar:

- Text is antialiased or otherwise rendered cleanly enough for report screenshots.
- Font size, weight, color, and baseline placement look intentional.
- Cell alignment, padding, wrapping, clipping, and merged-cell layout are visually coherent.
- Borders, gridlines, fills, and backgrounds match Office-like proportions.
- Charts look like polished chart exports, with legible titles, axes, labels, legends, and series geometry.
- Images preserve aspect, transparency, placement, and clipping where supported.
- SVG and PNG are visually comparable for the same source.
- Unsupported effects are diagnosed, but supported output must not look rough by default.

Use automated tests for contracts and dimensions, but require human visual review or approved visual baselines for renderer changes.

## Naming, Namespaces, And Packages

### Package Plan

Do not create a new NuGet package for the first Excel range-to-image work.

Use the packages that already express the correct ownership:

- `OfficeIMO.Drawing`
  - Owns shared visual primitives, SVG export, raster buffers, raster canvas, PNG encoding, and document-agnostic image diagnostics.
  - Remains dependency-free.
- `OfficeIMO.Excel`
  - Owns Excel range, worksheet, and workbook visual snapshots.
  - Owns the friendly Excel APIs for range, worksheet, and workbook image export.
  - Uses `OfficeIMO.Drawing` for rendering and encoding, which it already references.
- `OfficeIMO.Excel.Pdf`
  - Keeps PDF-specific export behavior.
  - Reuses the neutral Excel visual snapshot once it is extracted, but should not own the snapshot model.
- `OfficeIMO.PowerPoint`, `OfficeIMO.Word`, and other document packages
  - Later own their document-specific visual snapshots and thin image export APIs.
  - Reuse `OfficeIMO.Drawing` instead of creating their own raster stack.

Reserve a new shared package only if the feature family later outgrows `OfficeIMO.Drawing`. If that happens, prefer `OfficeIMO.Imaging` for cross-document orchestration and image export services. Do not start with `OfficeIMO.Imaging`, because the first missing engine is drawing/raster functionality and the repo already has `OfficeIMO.Drawing` for that role.

Avoid `OfficeIMO.Excel.Image` for the first implementation. It would add package friction without removing dependencies, because Excel already depends on Drawing. A separate `OfficeIMO.Excel.Image` package should only be reconsidered if image export becomes large enough to justify optional install size or release cadence separation.

### Namespace Plan

Follow existing OfficeIMO style: public document types mostly live in the document namespace, with folders used for organization.

Recommended public namespaces:

- `OfficeIMO.Drawing`
  - Public shared image primitives and diagnostics.
  - Public raster types only when users reasonably need them.
- `OfficeIMO.Excel`
  - Public Excel image export options, results, and extension methods.
  - Excel visual snapshots if and when they become public contracts.
- `OfficeIMO.PowerPoint`
  - Later PowerPoint slide image options and APIs.
- `OfficeIMO.Word`
  - Later Word page image options and APIs.

Recommended internal organization:

- `OfficeIMO.Drawing/Raster/`
  - `OfficeRasterImage`
  - `OfficeRasterCanvas`
  - `OfficeRasterColor`
  - `OfficeAlphaBlend`
- `OfficeIMO.Drawing/Png/`
  - `OfficePngWriter`
  - PNG filtering, CRC, and chunk helpers.
- `OfficeIMO.Drawing/Rendering/`
  - `OfficeDrawingRasterRenderer`
  - shared text and shape rendering helpers.
- `OfficeIMO.Excel/Imaging/`
  - `ExcelRangeVisualSnapshot`
  - `ExcelWorksheetVisualSnapshot`
  - `ExcelImageExportOptions`
  - `ExcelRangeImageRenderer`
  - `ExcelSheet.Imaging.cs`
  - `ExcelRange.Imaging.cs` if a first-class range type exists.

The folder names may be specific, but namespaces should stay simple unless a public surface becomes large enough to deserve a subnamespace. For example, prefer `namespace OfficeIMO.Excel` for `ExcelImageExportOptions` over `namespace OfficeIMO.Excel.Imaging` unless the API set becomes broad enough that a separate using is helpful.

### Type Naming

Use `Image` for user-facing concepts and `Raster` only for pixel-buffer implementation details.

Preferred shared names:

- `OfficeImageExportFormat`
  - Values: `Png`, `Svg`.
- `OfficeImageExportResult`
  - Format, dimensions, bytes or output metadata, diagnostics.
- `OfficeImageExportDiagnostic`
  - Severity, code, message, source reference where available.
- `OfficeImageExportDiagnosticSeverity`
  - `Info`, `Warning`, `Error`.
- `OfficeRasterImage`
  - Internal or advanced public RGBA buffer.
- `OfficeRasterCanvas`
  - Internal or advanced public pixel drawing canvas.
- `OfficePngWriter`
  - Low-level PNG encoder.

Preferred Excel names:

- `ExcelImageExportOptions`
  - Shared Excel image options such as scale, gridlines, include images, include charts, and diagnostics behavior.
- `ExcelRangeImageExportOptions`
  - Add only if range export needs options that do not belong to worksheet/workbook export.
- `ExcelWorksheetImageExportOptions`
  - Used when worksheet export needs print area, used range, page slicing, or page setup behavior.
- `ExcelWorkbookImageExportOptions`
  - Used when workbook export needs sheet selection and output naming.
- `ExcelRangeVisualSnapshot`
  - Format-neutral visual model for a selected range.
- `ExcelWorksheetVisualSnapshot`
  - Format-neutral visual model for a worksheet or page slice.
- `ExcelImageExportResult`
  - Excel-specific result if generic `OfficeImageExportResult` is not enough.

Avoid these names:

- `AllToImage`, `OfficeToImage`, or `DocumentToImage`
  - Too broad before Word, PowerPoint, PDF, and Excel share proven contracts.
- `Screenshot`
  - Implies desktop Office/browser fidelity and behavior.
- `Bitmap`
  - Too narrow and platform-loaded; prefer `Raster`.
- `Convert`
  - Too vague for APIs. Prefer `ToPng`, `ToSvg`, `SaveAsPng`, `SaveAsSvg`, and `ExportImages`.
- `Renderer` in public user-facing APIs
  - Keep renderers internal until custom renderer injection is a proven need.

### API Naming

Use format-specific convenience methods for the simple cases and result-returning export methods for advanced cases:

```csharp
byte[] png = sheet.Range("A1:D12").ToPng();
string svg = sheet.Range("A1:D12").ToSvg();

sheet.Range("A1:D12").SaveAsPng("range.png");
sheet.Range("A1:D12").SaveAsSvg("range.svg");

OfficeImageExportResult result = sheet.Range("A1:D12").ExportImage(
    OfficeImageExportFormat.Png,
    new ExcelImageExportOptions { Scale = 2 });
```

For multi-output operations, avoid pretending there is one image:

```csharp
IReadOnlyList<OfficeImageExportResult> pages = sheet.ExportImages(options);
IReadOnlyList<OfficeImageExportResult> images = workbook.ExportImages(options);
```

This gives the small first feature friendly names while keeping room for worksheet pages, workbook sheet collections, and later document families.

## Phase Plan

### Phase C0: Renderer Brain Consolidation

Purpose: make sure OfficeIMO has one reusable image rendering engine before premium Excel work deepens a parallel implementation.

Deliverables:

- Inventory every product PNG/SVG/raster/image-export path and classify it as shared engine, document adapter, PDF-specific writer behavior, test-only helper, or duplicate private renderer.
- Promote reusable Visio PNG renderer capabilities into `OfficeIMO.Drawing`: dashed strokes, solid elliptical arcs, rotated image drawing, rotated ellipse fill/stroke, even-odd contour/path fills, text measurement/layout helpers, clipping, supersampling/downsampling, PNG decode/encode reuse, and reusable visual diagnostics where appropriate. PNG decode/encode, supersampled pixel storage/blending/resolve, polygon/contour fills, line/polyline/dashed strokes, solid elliptical arcs, rotated ellipse fill/stroke, rotated/scaled image projection, anchored text-line drawing, fallback glyph drawing, text measurement, shared text wrapping/trim helpers, and rectangular raster clipping are already in shared Drawing on this branch.
- Migrate `OfficeIMO.Visio` native PNG export to consume shared Drawing raster/PNG primitives while keeping Visio-specific page, shape, connector, stencil, and label-layout semantics in `OfficeIMO.Visio`.
- Move or wrap PDF/Visio visual-baseline PNG comparison helpers into shared test support. This branch now has shared support for Excel, Visio, and PDF raster baseline comparisons.
- Update docs to forbid future document-specific private pixel engines unless there is a documented, narrow source-format reason.

Acceptance:

- `OfficeIMO.Drawing` owns the reusable raster/PNG/text/path primitives used by Excel and Visio, including measured line layout primitives for plain multiline text.
- `OfficeIMO.Visio` no longer has a private PNG encoder/decoder or private supersampled pixel buffer/resolve loop, and its remaining general-purpose private raster canvas is either removed or demonstrably only document-specific layout glue over shared Drawing primitives.
- Excel, Visio, and PDF visual-baseline raster comparisons use shared test support instead of private PNG decoder/encoder copies.
- Existing Visio PNG/SVG tests and premium native visual baselines pass or have intentional, reviewed baseline updates.
- Excel range/worksheet/workbook image tests still pass through the same shared Drawing engine.
- No dependency is added.

### Phase 0: Contracts And Extraction Seams

Purpose: make the small first feature fit the big goal.

Deliverables:

- Decide the public naming pattern for image export APIs, options, results, and diagnostics.
- Add internal contracts for visual snapshots and image export results.
- Extract reusable Excel visual planning from `OfficeIMO.Excel.Pdf` into neutral Excel-owned types without changing PDF behavior.
- Keep `OfficeIMO.Excel.Pdf` consuming the same data it consumes today, but through the neutral layer.
- Document the expected PNG/SVG contract and non-goals.

Acceptance:

- Existing Excel-to-PDF tests still pass.
- No runtime dependencies are added.
- The extracted model is format-neutral and does not mention PDF, PNG, or SVG in core names.

### Phase 1: Drawing PNG Foundation

Purpose: create the dependency-free pixel engine once.

Deliverables:

- Add a small RGBA image buffer.
- Add a canvas capable of fills, lines, rectangles, clipping, alpha blending, and basic text.
- Add a PNG encoder.
- Add `OfficeDrawing` to PNG export.
- Keep existing `OfficeDrawing` to SVG behavior intact.

Acceptance:

- PNG output has valid signature and dimensions.
- Basic drawing primitives produce nonblank images.
- SVG and PNG render the same simple drawing scenarios at the contract level.
- No product code uses external graphics libraries.

### Phase 2: Excel Range Visual Snapshot

Purpose: describe an Excel range visually without choosing an output format.

Deliverables:

- Add `ExcelRangeVisualSnapshot` or equivalent.
- Add snapshot options for scale, gridlines, hidden rows/columns, merged cells, images, charts, and style inclusion.
- Reuse existing worksheet reading, style snapshots, row/column metadata, merged ranges, images, and chart snapshots.
- Carry diagnostics for unsupported or approximated features.

Acceptance:

- Snapshot tests cover values, styles, borders, fills, dimensions, merges, hidden rows/columns, images, and charts where currently supported.
- Existing Excel PDF export remains behaviorally unchanged.

### Phase 3: Excel Range To SVG/PNG

Purpose: deliver the first user-visible feature.

Deliverables:

- Add range image APIs such as:

```csharp
byte[] png = sheet.Range("A1:D12").ToPng(options);
string svg = sheet.Range("A1:D12").ToSvg(options);
sheet.Range("A1:D12").SaveAsPng("range.png", options);
sheet.Range("A1:D12").SaveAsSvg("range.svg", options);
```

- Render cell backgrounds, text, gridlines, borders, merged cells, row heights, column widths, and simple alignments.
- Include worksheet images and chart snapshots when they intersect the range and can be represented.
- Return diagnostics for unsupported visuals.

Acceptance:

- Generated PNG/SVG artifacts exist, are nonblank, and have stable dimensions.
- Contract tests cover simple, styled, merged, hidden-row/column, image, and chart scenarios.
- API examples are small and match real entrypoints.

### Phase 4: Worksheet To Image

Purpose: expand from selected ranges to a complete sheet surface.

Deliverables:

- Add worksheet export options for used range, print area, explicit range, or page setup.
- Support one long worksheet image and page-sliced output.
- Respect sheet-level options such as gridlines, page orientation where relevant, print area, and scaling.
- Preserve the same renderer and diagnostics pipeline from range export.

Current progress:

- Worksheet export already reuses the range snapshot and shared renderer.
- `ExcelWorksheetImageExportOptions.Range` provides explicit range export.
- `ExcelWorksheetImageExportOptions.UsePrintArea` now uses the worksheet `_xlnm.Print_Area` defined name when configured.
- Explicit ranges override print areas.
- Missing print areas emit `ExcelPrintAreaMissing` and fall back to the used range.
- Single-image worksheet export keeps the legacy one-image contract: multi-area print areas emit `ExcelPrintAreaMultipleAreasUnsupported` and fall back to the used range.
- Multi-output worksheet export uses `ExcelSheet.ExportImages(...)` to split multi-area print areas into separate image results with `ExcelPrintAreaMultipleAreasSplit` diagnostics and `Sheet!Range` sources.
- `ExcelWorksheetImageExportOptions.SplitByManualPageBreaks` lets multi-output worksheet export split resolved ranges at manual row and column page breaks, honoring worksheet page order; single-image export emits `ExcelManualPageBreaksSingleImageUnsupported` instead of silently ignoring the request.
- Page-sliced image export now composes repeated print-title rows/columns through existing range snapshots, renders plain first/even/odd text header/footer chrome with supported page/sheet/workbook file/date/time fields in clipped and ellipsized left/center/right zones for multi-output PNG/SVG exports, and applies physical page orientation, margins, manual scale, supported paper-size geometry, and bounded one-page fit-to-width/fit-to-height scaling through shared image-layer composition. It still emits source-referenced `ExcelPageSetupUnsupported` diagnostics for automatic multi-page fit pagination requests, `ExcelPageSetupPaperSizeUnsupported` diagnostics for unknown paper codes, and `ExcelHeaderFooterUnsupported` diagnostics for richer header/footer images that are not rendered yet.

Acceptance:

- The worksheet path reuses range snapshot/rendering rather than duplicating it.
- Large sheets can be exported with bounded memory through tiling or page slicing where needed. The first page-slicing contract now covers manual row and column page breaks over explicit, used-range, and print-area selections, plus repeated print-title rows/columns for the multi-output path.
- Multi-page output returns a manifest/result collection, not only one byte array.

### Phase 5: Workbook To Images

Purpose: orchestrate all visible sheets.

Deliverables:

- Add workbook-level export options for selected sheets, visible sheets, print areas, and output naming.
- Return per-sheet and per-page results with diagnostics.
- Allow folder, stream factory, and in-memory outputs.

Current progress:

- Workbook image export is orchestration over worksheet export.
- `ExcelWorkbookImageExportOptions.SheetNames` selects sheets.
- `ExcelWorkbookImageExportOptions.UseWorksheetPrintAreas` forwards print-area intent to each worksheet, preserves per-sheet/per-area diagnostics, and now flattens multi-area worksheet output into the workbook result collection.
- `ExcelWorkbookImageExportOptions.SplitWorksheetsByManualPageBreaks` forwards manual page-break slicing to each worksheet and preserves page-order result ordering.
- Workbook options forward the shared worksheet visual switches instead of silently dropping drawing objects, conditional formatting, hyperlink hints, or comment bodies.
- Folder and in-memory output paths exist for the current sheet/result collection shape, with duplicate sheet filenames disambiguated when one sheet yields multiple images.

Acceptance:

- Workbook export is orchestration only.
- Sheet failures are isolated and reported.
- Results can be consumed by downstream tools without guessing filenames or page order.

### Phase 6: Shared Document Expansion

Purpose: reuse the same image stack outside Excel.

Likely order:

1. `OfficeDrawing` and chart PNG/SVG exports, because the drawing model already exists.
2. PowerPoint slide PNG/SVG, because slides are naturally bounded visual surfaces.
3. Word page PNG/SVG, after page layout contracts are explicit.
4. PDF page PNG/SVG, only when OfficeIMO has an internal PDF content rasterizer and does not depend on Poppler.

Acceptance:

- New document families produce visual snapshots first.
- PNG/SVG encoding remains shared.
- Document-specific packages stay thin.

### Phase 7: Unified Image Export Surface

Purpose: make the feature family feel consistent after several domains exist.

Deliverables:

- Align naming, options, result types, and diagnostics across Excel, Drawing, Word, PowerPoint, and PDF.
- Add shared helper abstractions only where repeated behavior has proven itself.
- Consider thin PowerShell/CLI wrappers after the .NET surface is stable.

Acceptance:

- Users can learn one export result/diagnostics model.
- Document-specific APIs remain discoverable and friendly.
- Shared abstractions reflect real reuse rather than speculative generality.

## Step-One Implementation Path

Initial implementation path:

1. Done initial slice: add dependency-free `OfficeRasterImage`, `OfficeRasterCanvas`, shared image export result/diagnostics, and `OfficePngWriter` in `OfficeIMO.Drawing`.
2. Done initial slice: add `ExcelRangeVisualSnapshot` and range snapshot building from values, styles, row/column metadata, merged ranges, and diagnostics.
3. Done initial slice: render Excel ranges to PNG and SVG with values, fills, gridlines, borders, row heights, column widths, and merged-cell coverage.
4. Done initial slice: add worksheet used-range/range export APIs over the same snapshot/renderer.
5. Done initial slice: add workbook `ExportImages` and `SaveAsImages` orchestration over worksheet exports.
6. Done initial slice: add embedded PNG worksheet image rendering and supported chart snapshot rendering through shared drawing primitives.
7. Done initial slice: add focused tests for shared PNG output plus Excel range, worksheet, workbook, embedded-image, and chart image export contracts.
8. Done first fidelity pass: add source-over alpha blending, antialiased raster primitives, bilinear image scaling, and PNG text alignment/style mapping.
9. Done consolidation slice: migrate shared Visio PNG internals for PNG read/write, supersampled storage/resolve, polygon and even-odd contour fills, line/dashed strokes, solid elliptical arcs, rotated ellipse fill/stroke, rotated/scaled image drawing, anchored text-line drawing, fallback glyph drawing, and text measurement into `OfficeIMO.Drawing`.
10. Done text fidelity slice: add wrapped Excel cell text, explicit/default vertical text alignment, SVG text clipping, and stable `ExcelCellTextClipped` diagnostics through a dedicated Excel image text-layout helper over shared Drawing measurement/rendering.
11. Done first baseline gate: add an approved Excel image PNG/SVG visual baseline with raster diff artifacts on mismatch, nonblank validation through shared PNG decode, and structural SVG assertions for text clipping, embedded images, charts, and percent display text.
12. Done visual QA consolidation slice: add shared Drawing-backed visual-baseline test support and migrate Excel image, Visio premium, and PDF raster baseline comparison paths to it.
13. Done text fidelity slice: capture styled cell font sizes and shrink-to-fit in Excel visual snapshots, expose thin cell/range APIs for them, and render them consistently through PNG/SVG with focused visual tests.
14. Done renderer consolidation slice: promote dashed ellipse stroke approximation into `OfficeIMO.Drawing`, keep Visio on the shared primitive, and add a Drawing raster contract test for dashed ellipse gaps.
15. Done text fidelity slice: capture Excel text rotation in visual snapshots, expose thin cell/range APIs, render basic numeric rotation and plain stacked rotation through the shared Drawing text renderer for PNG/SVG, and report approximation diagnostics with stable source references.
16. Done renderer consolidation slice: promote solid and dashed polyline stroking into `OfficeIMO.Drawing`, keep Visio connector/shape strokes on the shared primitive while preserving Visio's per-segment dash reset behavior, and prove it with focused Drawing tests plus the native Visio premium baseline gate.
17. Done image fidelity/diagnostics slice: detect worksheet image byte formats in Excel visual snapshots, embed known SVG-compatible formats such as JPEG in SVG output, and report unsupported PNG rasterization or SVG embedding with stable image-source diagnostics.
18. Done style fidelity slice: add a reusable Excel theme/indexed/direct color resolver with tint/shade support, wire it through `ExcelCell.GetStyle()`, inspection snapshots, and image visual snapshots, and prove PNG/SVG rendering with a theme-backed style contract test.
19. Done style fidelity slice: evaluate first-pass conditional formatting visuals in the neutral Excel image snapshot, render color scales and data bars to PNG/SVG, and report unsupported icon sets with stable source diagnostics.
20. Done visual QA slice: add an approved Excel conditional-formatting PNG/SVG baseline covering heat-map fills, positive and negative data bars, and unsupported icon-set diagnostics through the same Drawing-backed baseline comparison helper.
21. Done conditional-rule fidelity slice: expose differential fill colors in conditional rule snapshots, add optional fill colors to cell-is/formula rule authoring, render bounded numeric cell-is and simple comparison formula fills with priority and stop-if-true behavior, and extend the conditional-formatting baseline with a rule-driven fill column.
22. Done conditional diagnostics slice: emit stable source-referenced warnings for unsupported conditional rule types, unsupported data-bar/color-scale shapes, unsupported differential formats, text/non-numeric cell-is rules, and formula rules outside the simple numeric comparison subset.
23. Done hidden-layout contract slice: omit hidden rows/columns by default, honor `IncludeHidden`, and report hidden row/column omission plus hidden-anchored image/chart omission with stable source diagnostics.
24. Done rich-text fidelity slice: capture Excel rich text runs in the visual snapshot, render supported single-line runs to PNG/SVG with per-run bold/italic/underline/color/font-size mapping through shared Drawing text operations, include that path in the approved premium Excel visual baseline, and emit `ExcelCellRichTextLayoutApproximation` when rich text falls back to plain text because rotation would not preserve runs.
25. Done shared text measurement slice: add a bounded per-canvas cache to `OfficeRasterCanvas.MeasureText` so repeated Excel/Visio raster layout operations reuse deterministic Drawing-level measurements without global state.
26. Done shared raster clipping slice: add rectangular clip scopes to `OfficeRasterCanvas`, use them for Excel PNG cell text rendering so rotated/overflowing text cannot paint outside the cell bounds, and prove the primitive with focused Drawing tests plus an Excel rotated-text clipping contract.
27. Done visual QA performance slice: split Visio premium native visual baselines into per-scenario tests, expose `VisioPremiumGallery.CreateScenario(...)` for targeted scenario generation, and make Drawing raster coverage sampling adaptive for supersampled render targets so downsample antialiasing remains while redundant subpixel work is avoided.
28. Done style fidelity/diagnostics slice: carry Excel pattern fill metadata through `ExcelCell.GetStyle()`, inspection snapshots, and image visual snapshots; render pattern fills as deterministic hatch approximations in PNG/SVG through the shared Drawing primitives; and report `ExcelFillPatternApproximation` plus `ExcelFillGradientUnsupported` instead of silently flattening unsupported fill effects.
29. Done object diagnostics/consolidation slice: move worksheet comment and threaded-comment metadata resolution into a shared Excel utility used by inspection, feature reporting, and image export; report visible exported comments/notes and threaded comments with `ExcelCellCommentUnsupported` and `ExcelThreadedCommentUnsupported` source diagnostics instead of silently dropping them.
30. Done drawing-object diagnostics/consolidation slice: move worksheet drawing-object detection into a shared Excel utility used by PDF preflight and image export; report visible exported shapes, text boxes, connectors, group shapes, and non-chart graphic frames with `ExcelDrawingShapeUnsupported` instead of silently dropping them from image exports.
31. Done sparkline diagnostics/consolidation slice: move authored worksheet sparkline target-cell discovery into a shared Excel utility used by feature reporting, PDF preflight, and image export; report visible exported sparkline targets with `ExcelSparklineUnsupported` instead of silently dropping them from image exports.
32. Done sparkline rendering slice: carry visible same-sheet numeric sparklines into `ExcelRangeVisualSnapshot`, render line/column/win-loss sparklines to PNG/SVG through shared Drawing primitives, preserve basic authored colors/markers/axis/negative styling, and replace the blanket unsupported warning with `ExcelSparklineRenderingApproximation` plus specific `ExcelSparklineExternalRangeUnsupported`, `ExcelSparklineRangeUnsupported`, `ExcelSparklineKindUnsupported`, and `ExcelSparklineDataMissing` diagnostics for the cases still outside the renderer.
33. Done sparkline visual QA slice: add approved PNG/SVG sparkline visual baselines covering line, column, and win/loss sparklines, assert SVG structure/colors/clipping, validate approved PNG nonblank dimensions, and run them through the shared Drawing-backed visual-baseline comparison helper.
34. Done comment-indicator visual slice: carry visible comments/notes and threaded comments into `ExcelRangeVisualSnapshot`, render top-right cell indicators to PNG/SVG through shared Drawing primitives, keep unsupported-body diagnostics source-referenced, add decoded PNG/SVG object tests, and include the legacy comment marker in the approved premium Excel visual baseline.
35. Done image clipping visual slice: include worksheet images whose visual rectangle overlaps the selected range even if their anchor cell is outside it; preserve hidden-anchor omissions; render negative-position/clipped images to PNG/SVG; add focused decoded-pixel/SVG tests and a dedicated clipped-image approved visual baseline.
36. Done image transform visual slice: carry authored worksheet picture rotation into `ExcelRangeVisualSnapshot`, render basic rotated PNG images through shared Drawing rotated image sampling, emit SVG image rotation transforms, and add focused decoded-pixel/SVG tests plus a dedicated rotated-image approved visual baseline that was visually reviewed for readable layout.
37. Done image transform consolidation slice: replace separate scaled/cropped/rotated raster image loops with one shared Drawing projector that combines source rectangles, rotation, and horizontal/vertical flips; wire Excel crop-plus-flip-plus-rotation through it for PNG/SVG; remove stale flip and crop-plus-rotation unsupported diagnostics; add focused shared Drawing and Excel workbook tests plus a visually reviewed transformed-image baseline.
38. Done drawing-object rendering slice: classify worksheet drawing objects once, route simple rectangle/rounded-rectangle solid RGB shapes with plain text through the neutral Excel snapshot and shared `OfficeIMO.Drawing` PNG/SVG renderers, keep unsupported variants diagnosed, and add focused object tests plus a reviewed approved drawing-object baseline.
39. Done layered drawing-order slice: add source drawing order to Excel images, charts, and supported drawing objects; introduce an ordered `ExcelVisualDrawingLayer` overlay stream; render supported shapes/images/charts through one PNG/SVG dispatcher; and prove mixed shape/image order in both directions with snapshot order, SVG order, and decoded PNG pixel assertions.
40. Done text-layout consolidation slice: promote trim-to-width into `OfficeIMO.Drawing.OfficeTextLayoutEngine`, replace Excel's private wrapped-line and line type helpers with shared `OfficeTextLine` output, keep Excel-specific shrink/vertical/clipping/rich-text decisions in the adapter, and prove the shared contract plus public Excel multiline PNG/SVG output with focused tests.
41. Done Visio SVG text-layout consolidation slice: replace Visio SVG's private wrap/break/max-line measurement helpers with `OfficeTextLayoutEngine`, keep SVG-specific emission/alignment/background behavior in the Visio adapter, and prove hard-break SVG text, full `VisioSvgExport`, `VisioPngExport`, and native premium Visio baselines still pass.
42. Done text-block fit consolidation slice: promote wrapped text-block fit-down into `OfficeIMO.Drawing.OfficeTextLayoutEngine.FitWrappedText` and `OfficeTextBlockLayout`, replace Visio PNG/SVG private fit math with the shared helper, keep adapter-specific placement/emission/background/underline behavior in Visio, and prove the shared fit contract plus Visio PNG/SVG/premium baseline behavior.
43. Done text-placement consolidation slice: promote reusable horizontal anchor, measured-line-left, and vertical top placement into `OfficeIMO.Drawing.OfficeTextPlacement` plus `OfficeTextVerticalAlignment`; migrate Excel PNG/SVG text and rich-text placement plus Visio PNG/SVG text/background/underline placement onto the shared helper; keep document-specific alignment mapping in adapters; and prove Drawing placement, Excel image export, Visio SVG/PNG, and native premium Visio baselines.
44. Done text clipping consolidation slice: promote visible-height text-block clipping, last-visible-line ellipsis, and clipped-state reporting into `OfficeIMO.Drawing.OfficeTextLayoutEngine.ClipTextBlockToHeight` / `OfficeTextBlockLayout.Clipped`; replace Excel's private `ExcelTextLayout` result and max-line clipping helper with the shared block layout while keeping Excel-specific shrink-to-fit, rotation, rich-text fallback, and diagnostics in the adapter; and prove the shared clipping contract plus Excel image export.
45. Done shrink-to-fit consolidation slice: promote measured single-line font-size fitting into `OfficeIMO.Drawing.OfficeTextLayoutEngine.FitSingleLineFontSize`; replace Excel's private shrink-to-fit binary search with a thin policy wrapper that calls the shared helper; and prove already-fit, fitted, and minimum-floor behavior plus Excel image export.
46. Done bounded text-layout orchestration slice: promote generic bounded text block layout orchestration into `OfficeIMO.Drawing.OfficeTextLayoutEngine.LayoutTextBlock`; replace Excel's private cell text layout coordinator with the shared helper; keep Excel-specific wrap/shrink/rotation policy, rich-text fallback, vertical alignment, and diagnostics in the adapter; and prove shared shrink/wrap/clip/single-line behavior plus Excel image export.
47. Done rich text block layout slice: add shared `OfficeRichTextRun` / segment / line / block layout contracts and `OfficeTextLayoutEngine.LayoutRichTextBlock`; render Excel hard-break and wrapped rich text through the shared layout in PNG/SVG with per-run bold/italic/underline/color/font-size preservation; keep rotation fallback diagnosed; and prove the Drawing contract plus Excel PNG/SVG rich-text behavior.
48. Done rich text shrink-to-fit slice: add proportional run font-size scaling to shared `OfficeTextLayoutEngine.LayoutRichTextBlock`; render Excel shrink-to-fit rich text through the shared rich layout in PNG/SVG without rich-text approximation diagnostics; prove width fitting plus run style preservation through Drawing and Excel image export tests; and add a dedicated approved PNG/SVG rich-text baseline covering single-line, hard-break, wrapped, shrink-to-fit, and clipped rich text.
49. Done basic rotated rich text slice: preserve rich text runs for basic rotated Excel cell text in PNG/SVG instead of falling back to plain text; add shared Drawing point-rotation placement support; keep `ExcelCellTextRotationApproximation` diagnostics for the non-Excel-exact rotation path; and extend the rich-text approved baseline with a visually reviewed rotated styled run.
50. Done chart series-color slice: carry simple authored Excel chart series fill/line colors through `ExcelChartSnapshot` into shared `OfficeChartSeries.Color`, render those colors in PNG/SVG through the existing Drawing chart renderer, suppress the generic series-style approximation diagnostic for that supported simple color case, and prove SVG plus decoded PNG output with a focused Excel image-export test.
51. Done chart point/marker color slice: add authored Excel chart point-fill APIs, carry simple `c:dPt` solid fills plus simple marker fill/visibility through `ExcelChartSnapshot` into shared `OfficeChartSeries.PointColors` and marker flags, render those colors in PNG/SVG through the existing Drawing chart renderer, keep diagnostics for marker shape/size/outline and richer point styling, and prove snapshot state plus SVG and decoded PNG pixels with focused Excel image-export tests.
52. Done chart gridline style slice: carry simple authored Excel major-gridline color and gridline visibility through `ExcelChartSnapshot` into shared `OfficeChartStyle.GridLineColor` / `ShowGridLines`, render those settings in PNG/SVG through the existing Drawing chart renderer, keep `ExcelChartGridlineStyleApproximation` diagnostics for complex styling that the shared renderer does not honor yet, and prove snapshot state plus SVG and decoded PNG output with focused Excel image-export tests.
53. Done chart axis-line style slice: add authored Excel category/value axis-line APIs, carry simple solid axis-line color and no-line visibility through `ExcelChartSnapshot` into shared `OfficeChartStyle.AxisColor` and `OfficeChartLayout.ShowCategoryAxisLine` / `ShowValueAxisLine`, render those settings in PNG/SVG through the existing Drawing chart renderer, keep `ExcelChartAxisStyleApproximation` diagnostics for width/complex styling that the shared renderer does not honor yet, and prove snapshot state plus SVG and decoded PNG output with focused Excel image-export tests.
54. Done chart title-color slice: carry simple authored Excel chart title RGB text color into shared `OfficeChartStyle.TitleColor`, render it through the existing Drawing chart renderer, keep `ExcelChartTextStyleApproximation` diagnostics for title typography that the shared renderer does not honor yet, and prove snapshot state plus SVG and decoded PNG output with focused Excel image-export tests.
55. Done chart body/axis text-color slice: split chart text-style export helpers into a focused partial; carry simple shared legend/data-label text color into `OfficeChartStyle.TextColor`; carry simple shared axis label/title text color into `OfficeChartStyle.MutedTextColor`; keep `ExcelChartTextStyleApproximation` diagnostics for conflicting per-element colors, font size, bold, italic, and non-solid text fills that the shared chart renderer does not honor yet; and prove snapshot state plus SVG and decoded PNG output with focused Excel image-export tests.
56. Done chart axis number-format slice: carry simple authored Excel value-axis number formats into shared `OfficeChartLayout.VerticalAxisNumberFormat` for vertical value axes and `OfficeChartLayout.HorizontalAxisNumberFormat` for horizontal bar value axes; keep `ExcelChartAxisNumberFormatApproximation` diagnostics for date/time/text/conditional/scientific-style format shapes that the shared chart number formatter does not honor yet; and prove snapshot state plus SVG output with focused Excel image-export tests.
57. Done chart axis label-visibility slice: carry Excel `TickLabelPositionValues.None` for primary category/value axes into shared `OfficeChartLayout.ShowCategoryAxisLabels` and `ShowValueAxisLabels`, letting the shared chart renderer suppress labels instead of Excel keeping that behavior private; and prove chart-only SVG output plus snapshot flags with focused Excel image-export tests.
58. Done chart marker-size slice: carry authored Excel marker size through `ExcelChartSeries`, the neutral Excel chart snapshot, and shared `OfficeChartSeries.MarkerSize`; render sized line/scatter/radar markers through the shared Drawing chart renderer; stop treating simple marker size as a series-style approximation; and prove snapshot state plus SVG radius and decoded PNG pixels with focused Excel image-export tests.
59. Done chart marker-shape slice: add shared `OfficeChartMarkerShape`, carry authored Excel circle/square/diamond/triangle markers through the neutral and shared chart snapshots, render non-circle markers through Drawing rectangles/polygons in PNG/SVG, keep unsupported marker symbols diagnosed, and prove diamond marker shape with SVG polygon output plus decoded PNG pixels.
60. Done chart marker-outline slice: carry simple authored Excel marker solid outline color and width through `ExcelChartSeries`, the neutral Excel chart snapshot, and shared `OfficeChartSeries.MarkerOutlineColor` / `MarkerOutlineWidth`; render marker outlines in PNG/SVG through the shared Drawing chart renderer; stop treating simple marker outlines as a series-style approximation; and prove snapshot state plus SVG stroke and decoded PNG outline pixels with focused Excel image-export tests.
61. Done chart axis/gridline width slice: carry simple authored Excel axis-line and major-gridline widths through shared `OfficeChartStyle.AxisLineWidth` / `GridLineWidth`, render those widths in shared Drawing PNG/SVG output, stop treating simple width-only solid axis/gridline outlines as style approximations, and prove snapshot state plus SVG stroke widths and decoded PNG pixels with focused Excel image-export tests.
62. Done chart axis/gridline dash slice: carry simple Excel preset dashes for axis lines and major gridlines through shared `OfficeChartStyle.AxisLineDashStyle` / `GridLineDashStyle`, render those dashes through Drawing SVG and raster line rendering, stop treating simple preset dash axis/gridline outlines as style approximations, and prove snapshot state plus SVG dash arrays and decoded PNG pixels with focused Excel image-export tests.
63. Done chart series-line-width slice: carry simple authored Excel chart series line widths through `ExcelChartSeries`, the neutral Excel chart snapshot, and shared `OfficeChartSeries.StrokeWidth`; render stroked line/scatter/radar/area series with that width in shared Drawing output; stop treating simple series outline width as a series-style approximation; and prove the shared renderer plus Excel SVG/PNG output with focused contract tests.
64. Done chart series-line-dash slice: carry simple Excel preset dashes for chart series outlines through `ExcelChartSeries`, the neutral Excel chart snapshot, and shared `OfficeChartSeries.StrokeDashStyle`; render dashed line/scatter/area/radar series strokes through shared Drawing line primitives; stop treating simple preset series dashes as a series-style approximation; and prove the shared renderer plus Excel SVG/PNG output with focused contract tests.
65. Done chart marker-symbol slice: add shared plus and X marker shapes, render them as Drawing line primitives, carry authored Excel plus/X marker symbols through the neutral chart snapshot, stop treating those symbols as unsupported series-style approximations, and prove shared renderer plus Excel SVG/PNG output with focused tests.
66. Done chart marker-symbol completion slice: add shared dash, dot, and star marker shapes, render dash as a Drawing line primitive, dot as a centered Drawing ellipse, and star as the shared five-point DrawingML preset polygon, carry authored Excel dash/dot/star symbols through the neutral chart snapshot, keep picture markers diagnosed, and prove shared renderer plus Excel SVG/PNG output with focused tests.
67. Done Drawing geometry consolidation slice: add shared `OfficeGeometry` distance and polyline-by-length interpolation helpers, migrate Visio native PNG connector label placement, SVG connector label placement, and collision-aware label layout away from private interpolation copies, keep Visio-specific page-coordinate policy in Visio, and prove the shared geometry plus existing connector-label PNG/SVG contracts.
68. Done chart/plot area outline depth slice: carry simple authored Excel chart-area and plot-area outline widths plus preset dashes into shared `OfficeChartStyle`, render them in Drawing PNG/SVG output, keep richer area styling diagnosed with `ExcelChartAreaStyleApproximation`, and prove the shared renderer plus Excel SVG/PNG output with focused contract tests.
69. Done chart text font-size slice: carry simple authored Excel legend, data-label, and axis-label font sizes into shared `OfficeChartLayout`, render those sizes through the existing Drawing text renderer, keep chart title and font-family variants outside that slice diagnosed with `ExcelChartTextStyleApproximation`, and prove snapshot state plus SVG output with focused Excel image-export tests.
70. Done chart title typography slice: add shared `OfficeChartStyle.TitleFontSize` / `TitleFontStyle`, render simple authored Excel chart title font size plus bold/italic through the shared Drawing title renderer, keep richer text effects diagnosed with `ExcelChartTextStyleApproximation`, and prove both shared Drawing and Excel SVG output with focused contract tests.
71. Done chart title font-family slice: add shared `OfficeChartStyle.TitleFontFamily`, render simple authored Excel chart title font-family through the shared Drawing title renderer and SVG exporter, keep conflicting title font families and richer text effects diagnosed with `ExcelChartTextStyleApproximation`, and prove both shared Drawing and Excel SVG output with focused contract tests.
72. Done chart non-title font-family buckets: add shared `OfficeChartLayout.LegendFontFamily`, `DataLabelFontFamily`, and `AxisTextFontFamily`; render simple authored Excel legend, data-label, and axis-label font families through the shared Drawing text renderer and SVG exporter; keep conflicts inside each supported shared text bucket and richer text effects diagnosed with `ExcelChartTextStyleApproximation`; and prove shared Drawing plus Excel SVG output with focused contract tests.
73. Done chart non-title font-style buckets: add shared `OfficeChartLayout.LegendFontStyle`, `DataLabelFontStyle`, and `AxisTextFontStyle`; render simple authored Excel legend, data-label, and axis-label bold/italic through the shared Drawing text renderer and SVG exporter; keep conflicts inside each supported shared text bucket and richer text effects diagnosed with `ExcelChartTextStyleApproximation`; and prove shared Drawing plus Excel SVG output with focused contract tests.
74. Done chart axis-title font-size bucket: add shared `OfficeChartLayout.AxisTitleFontSize`; render simple authored Excel category/value axis-title font size through shared Drawing, including axis-title band sizing; keep conflicts inside the axis-title size bucket diagnosed with `ExcelChartTextStyleApproximation`; and prove shared Drawing plus Excel SVG output with focused contract tests.
75. Done chart axis-title font-family/style buckets: add shared `OfficeChartLayout.AxisTitleFontFamily` and `AxisTitleFontStyle`; render simple authored Excel category/value axis-title font family and bold/italic overrides separately from axis labels; keep conflicts inside the axis-label and axis-title buckets diagnosed with `ExcelChartTextStyleApproximation`; and prove shared Drawing plus Excel SVG output with focused contract tests.
76. Done chart axis-title text-color bucket: add shared `OfficeChartStyle.AxisTitleColor`; render simple authored Excel category/value axis-title text color separately from axis-label text color; keep conflicts inside the axis-label and axis-title color buckets diagnosed with `ExcelChartTextStyleApproximation`; and prove shared Drawing plus Excel SVG/PNG output with focused contract tests.
77. Done chart legend/data-label text-color buckets: add shared `OfficeChartStyle.LegendTextColor` and `DataLabelTextColor`; render simple authored Excel legend and data-label text colors separately instead of forcing them through one body text bucket; keep conflicts inside each supported body text bucket diagnosed with `ExcelChartTextStyleApproximation`; and prove shared Drawing plus Excel SVG/PNG output with focused contract tests.
78. Done chart category/value axis-line style buckets: add shared `OfficeChartStyle.CategoryAxisColor` / `ValueAxisColor`, `CategoryAxisLineWidth` / `ValueAxisLineWidth`, and `CategoryAxisLineDashStyle` / `ValueAxisLineDashStyle`; render simple authored Excel category and value axis-line colors, widths, and preset dashes separately for normal and horizontal bar chart orientation; keep the older `AxisColor` / `AxisLineWidth` / `AxisLineDashStyle` as shared fallbacks; and prove shared Drawing plus Excel SVG/PNG output with focused contract tests.
79. Done chart category/value major-gridline style buckets: add shared `OfficeChartStyle.CategoryGridLineColor` / `ValueGridLineColor`, `CategoryGridLineWidth` / `ValueGridLineWidth`, `CategoryGridLineDashStyle` / `ValueGridLineDashStyle`, and category/value major-gridline visibility overrides; render simple authored Excel category and value major-gridline colors, widths, and preset dashes separately for normal and horizontal bar chart orientation; keep the older `GridLineColor` / `GridLineWidth` / `GridLineDashStyle` / `ShowGridLines` as value-gridline fallbacks; and prove shared Drawing plus Excel SVG/PNG output with focused contract tests.
80. Done chart axis-placement diagnostics: initially report `ExcelChartAxisTickLabelPositionApproximation` for authored high/low tick-label placement, `ExcelChartAxisMinorTickMarkPlacementApproximation` for authored minor tick marks while placement remains approximate, and `ExcelChartAxisCrossingApproximation` for custom axis crossing, while keeping supported `none` tick-label suppression diagnostic-free; prove all three source-referenced warning codes with focused Excel image-export tests.
81. Done chart category/date axis number-format diagnostics: report `ExcelChartCategoryAxisNumberFormatUnsupported` for authored category or date axis number formats, because the shared chart renderer does not yet format category/date tick labels; keep simple rendered value-axis numeric formats diagnostic-free and keep complex rendered value-axis formats on `ExcelChartAxisNumberFormatApproximation`; prove the source-referenced warning with focused Excel image-export tests.
82. Done chart axis-scale diagnostics: report `ExcelChartAxisScaleApproximation` for authored log scale, unsupported value-axis or horizontal-bar reverse-order orientation, non-value-axis scale/unit settings, invalid value-axis min/max/unit values, and non-default cross-between settings; keep supported linear value-axis min/max/major/minor-unit scale diagnostic-free; prove source-referenced warnings for unsupported reverse-order scale with focused Excel image-export tests.
83. Done chart axis display-unit rendering slice: add shared `OfficeChartLayout` display-unit divisor/label buckets, carry authored Excel built-in and custom value-axis display units into the neutral chart snapshot, scale shared Drawing value-axis labels, render display-unit captions in PNG/SVG through the shared chart text path, and prove diagnostic-free SVG output with focused Excel image-export tests.
84. Done chart major tick-mark rendering slice: add shared `OfficeChartAxisTickMark`, carry authored Excel category/value major tick marks into `OfficeChartLayout`, render simple inside/outside/cross major axis ticks through shared Drawing, remove the major-tick unsupported diagnostic path, and leave exact minor tick-mark placement for a later chart-axis slice.
85. Done chart minor-gridline rendering slice: add shared `OfficeChartStyle` category/value minor-gridline color, width, dash, and visibility buckets; carry simple authored Excel minor gridlines into the neutral chart snapshot; render midpoint minor gridlines through shared Drawing behind major gridlines; and keep complex gridline effects diagnosed with `ExcelChartGridlineStyleApproximation`.
86. Done chart minor tick-mark rendering slice: carry authored Excel category/value minor tick marks into `OfficeChartLayout`, render simple inside/outside/cross minor ticks through shared Drawing with a stable `ExcelChartAxisMinorTickMarkPlacementApproximation` diagnostic, and prove both the shared renderer output and diagnostic contract with focused Excel image-export tests.
87. Done chart value-axis scale rendering slice: carry authored Excel linear value-axis minimum, maximum, major unit, and minor unit into `OfficeChartLayout`; apply the shared scale to plotted values, value-axis labels, major/minor gridlines, and major/minor tick marks for vertical and horizontal value axes; keep unsupported log/value-axis-reverse-order/non-value-axis-unit/cross-between variants diagnosed with `ExcelChartAxisScaleApproximation`; and prove the diagnostic-free SVG label and minor-gridline output with focused Excel image-export tests.
88. Done chart high/low tick-label placement slice: add shared `OfficeChartAxisTickLabelPosition`, carry authored Excel high/low/next-to/none tick-label positions into physical horizontal/vertical `OfficeChartLayout` axes, reserve plot space for high-side labels, render simple high-side vertical and horizontal axis labels through shared Drawing, and remove the high/low `ExcelChartAxisTickLabelPositionApproximation` warning with focused renderer and image-export tests.
89. Done chart maximum value-axis crossing slice: add shared `OfficeChartAxisCrossingPosition`, carry authored Excel value-axis `crosses=max` into physical vertical `OfficeChartLayout` axes for non-bar charts, render the vertical value axis and next-to labels on the right side through shared Drawing, and remove the `ExcelChartAxisCrossingApproximation` warning for that supported case with focused renderer and image-export tests.
90. Done chart maximum category-axis crossing slice: carry authored Excel category/date-axis `crosses=max` into physical horizontal `OfficeChartLayout` axes for non-bar charts, render the horizontal category axis and next-to labels above the plot through shared Drawing, reverse horizontal tick outside direction for top axes, and keep that supported case diagnostic-free with focused renderer and image-export tests.
91. Done chart category-axis reverse-order slice: add shared `OfficeChartLayout.ReverseCategoryAxis`, map authored Excel category/date-axis max-min orientation into the neutral layout for non-bar charts, render reversed category labels plus non-bar column/line/area category positions through shared Drawing, and remove `ExcelChartAxisScaleApproximation` for that supported case with focused renderer and image-export tests.
92. Done Visio PNG stroke-dash consolidation slice: add shared `OfficeRasterCanvas.DrawStyledPolyline`, `DrawPatternedPolyline`, `DrawStyledEllipse`, and `DrawPatternedEllipse`; map Visio line patterns into shared `OfficeStrokeDashStyle`; route Visio PNG polygon, polyline, connector, underline, database-shape, and ellipse strokes through shared Drawing instead of a private boolean-dashed adapter; and prove shared Drawing plus Visio PNG output with focused raster/export tests.
93. Done shared SVG stroke-dash consolidation slice: add shared `OfficeStrokeDashStyleExtensions.GetSvgDashArray`; route `OfficeDrawingSvgExporter`, Excel range SVG border rendering, and Visio SVG shape/connector stroke output through the shared formatter; move Visio line-pattern mapping into one internal mapper reused by PNG and SVG; and prove Drawing, Visio SVG/PNG, and Excel image-export contracts with focused and broad tests.
94. Done shared SVG formatting consolidation slice: add `OfficeSvgFormatting` for invariant SVG number formatting, XML escaping, CSS RGB color formatting, alpha-to-opacity conversion, and writer color attributes; route `OfficeDrawingSvgExporter`, Excel range SVG number/text escaping, Visio SVG numeric formatting, and Visio SVG color/opacity writing through the shared Drawing helper; and prove Drawing, Visio SVG, and Excel image-export contracts.
95. Done shared SVG `StringBuilder` attribute-emission slice: add reusable `OfficeSvgFormatting.AppendAttribute`, `AppendNumberAttribute`, and `AppendPaintAttribute`; route Excel range SVG root/background/grid/data-bar, border/pattern line, pattern-fill rectangle, and plain/rich text-start emission through the shared Drawing helper so number formatting, XML escaping, CSS RGB colors, and alpha opacity are no longer Excel-private in those central paths; and prove Drawing plus Excel image-export contracts.
96. Done shared SVG clip/point-list consolidation slice: add reusable `OfficeSvgFormatting.AppendClipPathReference`, `AppendRectClipPathDefinition`, and `AppendPointsAttribute`; route Excel SVG image clip paths, cropped/transformed image references, sparkline clip groups, sparkline polyline/axis/bar/marker output, drawing-object SVG shells, and comment-indicator polygons through shared Drawing formatting helpers; and prove Drawing plus Excel image-export contracts.
97. Done shared SVG clip/rotation completion slice: add reusable `OfficeSvgFormatting.FormatRotateTransform`, `AppendRotateTransformAttribute`, and `WriteRotateTransformAttribute`; route Excel SVG text, rich-text, pattern-fill clip groups, simple rotated worksheet images, and Visio SVG shape/text rotation through shared Drawing formatting helpers so clip-path and rotate-transform assembly no longer lives in multiple renderer brains; and prove Drawing, Excel image-export, and Visio SVG contracts.
98. Done shared SVG matrix-transform consolidation slice: add reusable `OfficeSvgFormatting.FormatMatrixTransform` and `AppendMatrixTransformAttribute`; route `OfficeDrawingSvgExporter` shape placement/local-coordinate matrix transforms and clip-path group references through shared SVG formatting helpers instead of private transform/clip string assembly; and prove Drawing exporter, Excel image-export, and Visio SVG contracts.
99. Done shared SVG stroke cap/join consolidation slice: add reusable `OfficeSvgFormatting.FormatStrokeLineCap`, `FormatStrokeLineJoin`, `AppendStrokeLineCapAttribute`, `AppendStrokeLineJoinAttribute`, `WriteStrokeLineCapAttribute`, and `WriteStrokeLineJoinAttribute`; route `OfficeDrawingSvgExporter`, Excel dotted-border SVG output, Visio connector SVG output, and Visio stencil primitive SVG output through shared Drawing SVG stroke helpers instead of private enum mapping or repeated literal `round` attributes; and prove Drawing, Excel image-export, and Visio SVG contracts.
100. Done shared SVG stroke dash-array consolidation slice: add reusable `OfficeSvgFormatting.AppendStrokeDashArrayAttribute`, `AppendStrokeDashStyleAttribute`, `WriteStrokeDashArrayAttribute`, and `WriteStrokeDashStyleAttribute`; route `OfficeDrawingSvgExporter`, Excel styled border SVG output, Visio connector SVG output, and Visio shape SVG output through shared Drawing SVG dash helpers instead of repeated `stroke-dasharray` attribute emission; and prove Drawing, Excel image-export, and Visio SVG contracts.
101. Done worksheet/workbook page-output slice: add worksheet `ExportImages(...)` for multi-output image export, split multi-area print areas into separate PNG/SVG-capable results with stable diagnostics, route workbook export through worksheet multi-output orchestration, and keep saved multi-area outputs distinct on disk.
102. Done manual page-break image-output slice: add worksheet and workbook options for manual row/column page-break splitting, preserve the single-image warning contract, honor worksheet page order, and prove explicit-range plus workbook multi-result output through decoded PNG-backed tests.
103. Done page-chrome diagnostics slice: report unsupported print-title repetition, physical page setup orientation/scaling, and header/footer chrome when callers request page-sliced image export, keeping those premium worksheet/page gaps explicit instead of silently dropping them.
101. Done shared SVG writer numeric-attribute consolidation slice: add reusable `OfficeSvgFormatting.WriteNumberAttribute` and `WriteViewBoxAttribute`; route Visio SVG root dimensions/viewBox, background rectangles, stencil primitive coordinates, connector stroke widths, shape ellipse/image geometry, shape stroke widths, and text/tspan numeric placement through shared Drawing SVG writer formatting instead of per-call `Format(...)` attributes; and prove Drawing plus Visio SVG contracts.
102. Done shared SVG move-line path-data consolidation slice: add reusable `OfficeSvgFormatting.FormatMoveLinePathData` and `AppendMoveLinePathData` for invariant `M`/`L`/`Z` SVG path serialization; route Visio SVG open connector paths and closed shape/preserved-geometry paths through the shared Drawing formatter while keeping Visio page-coordinate conversion in the Visio adapter; and prove Drawing plus Visio SVG contracts.
103. Done shared Drawing SVG primitive-geometry consolidation slice: route `OfficeDrawingSvgExporter` rectangle, rounded-rectangle, ellipse, and line numeric geometry attributes through `OfficeSvgFormatting.AppendNumberAttribute`, route polygon point-list output through `AppendPointsAttribute`, and keep the exporter as the central shared Drawing SVG surface instead of preserving raw per-primitive number/point formatting; prove the shared Drawing SVG exporter contracts.
104. Done shared SVG path-command data consolidation slice: add reusable `OfficeSvgFormatting.FormatPathData` and `AppendPathData` for shared `OfficePathCommand` SVG `d` serialization with optional offsets; route `OfficeDrawingSvgExporter` shape path output and clip-path path output through the shared formatter instead of duplicate path-command switch blocks; and prove the shared Drawing SVG formatter/exporter contracts.
105. Done shared Visio arrowhead path-data consolidation slice: route connector arrowhead SVG triangle paths through `OfficeSvgFormatting.FormatMoveLinePathData` instead of preserving a manual Visio-only `M/L/Z` string builder, keeping arrowhead geometry local while sharing the SVG path serialization brain.
106. Done Excel image diagnostic-code contract slice: add public `ExcelImageExportDiagnosticCodes` constants for current stable Excel image export diagnostics so callers can filter unsupported/approximate rendering without copying magic strings; route text and fill diagnostics (`ExcelCellTextClipped`, text rotation/stacked/rich-text approximation codes, and pattern/gradient fill codes) through the constants; and update focused Excel image tests to consume the same contract.
107. Done Excel image diagnostic-code source-of-truth routing slice: route all current Excel image export product diagnostic emissions through `ExcelImageExportDiagnosticCodes`, including chart approximation/unsupported codes, conditional-formatting unsupported codes, image format/anchor/decode codes, print-area fallback codes, hidden row/column omission codes, comment/threaded-comment unsupported codes, sparkline unsupported/approximation codes, and drawing-object unsupported/hidden-anchor codes; raw Excel image diagnostic-code strings now live in the constants surface instead of scattered renderer branches.
108. Done opt-in comment body rendering slice: add `ExcelImageExportOptions.ShowCommentBodies`, carry visible classic/threaded comment body payloads and cell-side anchor points through `ExcelRangeVisualSnapshot`, render first-pass dependency-free callout bodies with anchored pointers in PNG/SVG using shared Drawing shapes and `OfficeTextLayoutEngine`, route enabled bodies through the ordered `ExcelVisualDrawingLayer` stream instead of a separate post-pass renderer, change enabled-body diagnostics from unsupported to stable `ExcelCellCommentBodyApproximation` / `ExcelThreadedCommentBodyApproximation` codes, and prove the behavior with focused decoded-PNG/SVG tests plus a manually reviewed QA artifact.
109. Done shared dash vocabulary slice: add `OfficeStrokeDashStyleMapper` to dependency-free `OfficeIMO.Drawing`, move Visio `LinePattern` rendering away from the private `VisioLinePatternMapper`, route Excel chart preset-dash mapping through the same shared mapper without adding OpenXML dependencies to Drawing, and prove Visio integer patterns plus Office preset dash names with Drawing-layer contract tests.
110. Done Visio SVG path-command consolidation slice: add shared `OfficePathCommand.QuadraticBezierTo` and `OfficeSvgFormatting` `Q` path serialization; route supported built-in Visio SVG cylinder, shield, hexagon, cloud, person, monitoring, and database geometry paths through `OfficeSvgFormatting.FormatPathData` / `FormatMoveLinePathData` and shared `OfficePathCommand` instead of preserving local `M`/`L`/`Q`/`C`/`Z` string assembly; and prove Drawing plus Visio SVG still build/tests through the focused export contracts.
111. Done raster path-command fidelity slice: add shared Drawing path flattening for line, quadratic, cubic, and closed contours; route `OfficeDrawingRasterRenderer` path output through the flattener so PNG rendering follows Bezier geometry instead of drawing endpoint-only chords; fill closed contours through the shared even-odd raster fill path; stroke open and closed contours through shared styled polylines; and prove curved raster output with focused Drawing tests.
112. Done custom number-format display slice: extend the shared Excel image/autofit number-format helper so custom literal affixes, escaped literal characters, and positive/negative/zero format sections show in image snapshots and SVG output instead of falling back to raw values or stripped numbers.
113. Done simple gradient fill slice: add shared Drawing raster linear-gradient rectangle fills and shared SVG gradient-definition emission; resolve Excel two-stop linear gradient cell fills through one utility shared by `GetStyle()` and inspection snapshots; render them in Excel PNG/SVG output; keep unresolved/path/multi-stop gradients source-diagnosed.
114. Done shared text-block renderer slice: add `OfficeTextBlockRenderer` in dependency-free Drawing for measured plain text-block PNG/SVG emission with alignment, vertical placement, underline, rotation, and shared SVG style attributes; route non-rotated Excel plain cell text PNG/SVG output and Visio native PNG text output through that helper while keeping Excel diagnostics/rich text policy and Visio label-background/page-coordinate policy in their thin adapters; prove the shared renderer with Drawing contract tests plus focused Excel/Visio export tests.
115. Done shared SVG text-block writer slice: extend `OfficeTextBlockRenderer` with an `XmlWriter` text/tspan writer for measured text blocks, including shared font, fill/opacity, text-anchor, dominant-baseline, underline, bold/italic, rotation, and adapter-supplied attributes; route Visio SVG shape text and connector labels through that helper while keeping Visio-owned background rectangles, label-adjusted markers, coordinate mapping, and style resolution in the Visio adapter; prove the helper and migrated consumer with Drawing SVG-writer tests plus Visio SVG text/label contracts.
116. Done comment-body text renderer consolidation slice: route opt-in Excel comment/threaded-comment callout body text through `OfficeTextBlockRenderer.DrawRasterTextBlock` and `AppendSvgTextBlock` instead of local per-line PNG/SVG loops; keep Excel-owned title placement, callout geometry, pointers, drawing-layer routing, and source diagnostics in the Excel adapter; prove existing comment-body PNG/SVG output and diagnostics still pass through focused object tests.
117. Done shared hatch-pattern primitive slice: add neutral `OfficeHatchPatternKind`, shared raster `OfficeRasterCanvas.DrawHatchPatternRectangle`, and shared SVG `OfficeSvgFormatting.AppendHatchPatternRectangle`; route Excel pattern-fill PNG/SVG hatch output through those Drawing primitives while keeping OpenXML pattern-name mapping, density policy, and `FillPatternApproximation` diagnostics in the Excel adapter; prove Drawing contracts plus a dedicated approved Excel pattern-fill PNG/SVG baseline that was visually reviewed.
118. Done shared sparkline renderer slice: add neutral `OfficeSparklineKind`, `OfficeSparklineStyle`, `OfficeSparklinePointStyle`, and `OfficeSparklineRenderer` for dependency-free line, column, and win/loss sparkline PNG/SVG geometry and emission; route Excel sparkline image output through that shared renderer while keeping OpenXML extraction, kind mapping, per-point color and marker policy, approximation diagnostics, clipping, and source references in the Excel adapter; prove Drawing contracts and the existing approved Excel sparkline baseline without regenerating it.
119. Done shared data-bar renderer slice: add `OfficeDataBarRenderer` for dependency-free resolved proportional data-bar PNG/SVG output; route Excel conditional-formatting data-bar painting through the shared primitive while keeping rule evaluation, start/width ratios, colors, unsupported icon-set diagnostics, and source references in the Excel adapter; prove Drawing contracts plus the existing approved conditional-formatting baseline without regenerating it.
120. Done shared SVG image projector slice: add `OfficeSvgImageRenderer` for dependency-free SVG image projection with normalized source crop, clip rectangles, rotation, horizontal/vertical flips, and data URI construction; route Excel worksheet image SVG output through that shared primitive while keeping OpenXML anchor/crop/transform extraction, content-type allow-listing, and source diagnostics in the Excel adapter; prove Drawing contracts plus approved clipped-image, two-cell image, cropped-image, rotated-image, and transformed-image Excel baselines, then manually review representative image artifacts.
121. Done shared SVG image writer reuse slice: extend `OfficeSvgImageRenderer` with an `XmlWriter` image emitter for dependency-free SVG image output with shared number formatting, data URI reuse, preserve-aspect support, rotation, and flips; route Visio package-preview SVG artwork through that writer while keeping Visio-owned preview discovery, package metadata sniffing, placement, and shape-coordinate policy in the Visio adapter; prove Drawing writer output plus existing Visio package-preview SVG contracts for PNG projection, generic metadata sniffing, content-type parameter normalization, unsafe SVG fallback, and rotation.
122. Done shared SVG primitive writer slice: add `OfficeSvgPrimitiveWriter` for dependency-free `XmlWriter` circle, rectangle, line, and path emission with shared number/color/stroke-cap/stroke-join handling; route Visio built-in stencil SVG artwork through it while keeping stencil semantics and placement in Visio; and prove Drawing primitive output plus Visio stencil metadata/rotation contracts.
123. Done shared nested SVG wrapper slice: add `OfficeSvgFormatting.ExtractSvgInner`, `AppendNestedSvgStart`, `AppendNestedSvgEnd`, and `AppendNestedSvg`; route Excel chart, drawing-object, and comment-body SVG wrapper emission through the shared helper while keeping Excel-owned visual policy in the adapter; and prove Drawing formatting plus Excel chart/object/comment and approved baseline contracts.
124. Done shared SVG polygon element slice: add `OfficeSvgFormatting.AppendPolygonElement` overloads for complete dependency-free SVG polygon emission; route `OfficeDrawingSvgExporter` polygon shapes and Excel comment indicator/body-pointer SVG polygons through the shared helper while keeping geometry and source policy in the adapters; and prove Drawing formatter/exporter plus Excel comment/baseline contracts.
125. Done shared SVG line element slice: add `OfficeSvgFormatting.AppendLineElement` overloads for complete dependency-free SVG line emission; route `OfficeDrawingSvgExporter` line shapes and Excel border SVG lines through the shared helper while keeping Drawing transform policy and Excel border-style policy in the adapters; and prove Drawing formatter/exporter plus Excel border/baseline contracts.
126. Done shared SVG rectangle element slice: add `OfficeSvgFormatting.AppendRectElement` overloads for complete dependency-free SVG rectangle and rounded-rectangle emission; route `OfficeDrawingSvgExporter` rectangle shapes, Excel gridline/cell-fill SVG rectangles, shared data-bar rectangles, and shared sparkline column/win-loss rectangles through the shared helper while keeping source style and transform policy in the adapters; and prove Drawing formatter/exporter/data-bar/sparkline plus Excel pattern-fill/conditional-formatting/sparkline/border/premium baseline contracts.
127. Done shared SVG sparkline polyline/circle slice: add `OfficeSvgFormatting.AppendPolylineElement` and `AppendCircleElement` overloads for complete dependency-free SVG polyline and circle emission; route `OfficeSparklineRenderer` line-series SVG polylines and marker circles through the shared helpers while preserving approved sparkline SVG attribute order and renderer-owned scaling/color policy; and prove Drawing formatter/sparkline plus Excel sparkline/premium baseline contracts.
128. Done shared SVG ellipse element slice: add `OfficeSvgFormatting.AppendEllipseElement` overloads for complete dependency-free SVG ellipse emission; route `OfficeDrawingSvgExporter` ellipse shapes through the shared helper while keeping Drawing placement, paint, and transform policy in the exporter; and prove Drawing formatter/exporter contracts.
129. Done shared SVG path element slice: add `OfficeSvgFormatting.AppendPathElement` overloads for complete dependency-free SVG path emission; route `OfficeDrawingSvgExporter` path shapes and path clip definitions through the shared helper while keeping Drawing placement, clip, paint, and transform policy in the exporter; and prove Drawing formatter/exporter contracts.
130. Done shared SVG clip-rectangle slice: route `OfficeDrawingSvgExporter` rectangle and rounded-rectangle clip-path definitions through `OfficeSvgFormatting.AppendRectElement` while keeping Drawing clip semantics in the exporter; and prove Drawing clip-path/exporter contracts.
131. Done shared SVG positioned-text slice: add `OfficeTextBlockRenderer.AppendSvgTextElement` for complete positioned SVG text/tspan emission; route `OfficeDrawingSvgExporter` drawing text boxes through the shared helper while keeping Drawing text-box placement and font/style policy in the exporter; and prove Drawing text renderer/exporter contracts, including shared fill opacity.
132. Done shared Excel rotated SVG text slice: extend `OfficeTextBlockRenderer.AppendSvgTextElement` with underline and rotation support; route Excel's plain rotated cell-text SVG output through the shared Drawing helper while keeping Excel rotation, clipping, alignment, style/color resolution, and diagnostics in the adapter; and prove Drawing helper plus public Excel rotated PNG/SVG text contracts.
133. Done shared Excel rich-text SVG segment slice: add `OfficeTextBlockRenderer.AppendSvgRichTextSegment`; route Excel rich cell text SVG segment emission through the shared Drawing helper while keeping run extraction, style fallback, line layout, cursor placement, clipping, rotation grouping, and diagnostics in Excel; and prove Drawing helper plus public Excel rich-text SVG contracts.
134. Done shared Excel comment-title SVG text slice: route comment-body title text through `OfficeTextBlockRenderer.AppendSvgTextElement` while keeping Excel comment-body geometry, clipping, colors, source references, and approximation diagnostics in the adapter; and prove the public comment-body SVG/PNG object-export contract still emits the title, bold style, start anchor, callout fill, and pointer.
135. Done shared Visio stencil-thumbnail SVG text slice: route generated stencil preview thumbnail captions through `OfficeTextBlockRenderer.AppendSvgTextElement` while keeping Visio gallery/package/thumbnail policy in the adapter; and prove the browser-renderable thumbnail artifact still emits the expected caption shape.
136. Done shared Excel SVG root-background rectangle slice: route the range SVG document background through `OfficeSvgFormatting.AppendRectElement` while keeping Excel canvas/viewBox/background policy in the adapter; and prove the public PNG/SVG export contract emits the expected root background rectangle.
137. Done shared Visio stencil-thumbnail wrapper SVG slice: extend `OfficeSvgImageRenderer.AppendImage` with optional `preserveAspectRatio`; route generated stencil preview thumbnail background/border rectangles, embedded preview image, and caption through Drawing helpers while keeping Visio thumbnail layout and package/gallery policy in the adapter; and prove the browser-renderable thumbnail artifact emits the expected wrapper shape.
138. Done Excel top/bottom conditional-formatting slice: expose top/bottom rank/bottom/percent metadata on rule snapshots, add fill-aware top/bottom builder overloads, render numeric top/bottom count and percent differential fills including ties, and emit `ExcelConditionalTopBottomUnsupported` only when the rule has no valid numeric candidates or rank; prove with snapshot, SVG, decoded PNG, and diagnostic assertions.
139. Done Excel duplicate-values conditional-formatting slice: add fill-aware sheet and fluent range duplicate-values APIs, render duplicate-values solid differential fills using nonblank visible cell values in the existing conditional fill pipeline, keep remaining unsupported unique-values/date/time/above-average rule families diagnosed, and prove with Open XML readback, snapshot, SVG, decoded PNG, and diagnostic assertions.
140. Done Excel unique-values conditional-formatting slice: add fill-aware sheet and fluent range unique-values APIs, render unique-values solid differential fills using the same distinctness helper as duplicate-values, keep remaining unsupported date/time/above-average rule families diagnosed, and prove with Open XML readback, snapshot, SVG, decoded PNG, and diagnostic assertions.
141. Done Excel above/below-average conditional-formatting slice: expose above/below/equal/std-dev metadata on rule snapshots, add fill-aware sheet and fluent range average APIs, render numeric above-average and below-average solid differential fills including equal-average variants, emit `ExcelConditionalAboveAverageUnsupported` for standard-deviation variants, and prove with Open XML readback, snapshot, SVG, decoded PNG, and diagnostic assertions.
142. Done Excel text conditional-formatting slice: expose text payload metadata on rule snapshots, add fill-aware sheet and fluent range contains/not-contains/begins-with/ends-with APIs, render case-insensitive text-rule solid differential fills, emit `ExcelConditionalTextRuleUnsupported` for malformed text rules, and prove with Open XML readback, snapshot, SVG, decoded PNG, and diagnostic assertions.
143. Done Excel print-title page-slice rendering slice: compose repeated print-title rows/columns around manual-page-sliced body ranges by reusing the existing Excel range snapshot renderer for each title/body component, stitch PNG through the shared raster canvas, stitch SVG through shared nested-SVG helpers, keep physical page setup and header/footer chrome diagnosed, and prove repeated title rows, repeated title columns, and row/column corner composition in page-sliced SVG/PNG output.
144. Done shared PDF PNG-container slice: add shared `OfficePngWriter` scanline/container entry points for dependency-free PNG signature, IHDR/IDAT/IEND chunk, CRC, and zlib writing; route PDF extracted-image PNG wrapping and soft-mask PNG output through those Drawing helpers while keeping PDF stream decoding and color-space policy in the PDF adapter; and prove Drawing-level scanline/container output plus existing PDF image-extractor contracts.
145. Done Excel page-sliced plain header/footer image slice: render simple odd-page left/center/right header and footer text around manual-page-sliced PNG/SVG outputs using shared Drawing text, raster, PNG, and nested-SVG helpers; preserve diagnostics for header/footer fields, formatting, and images; and prove SVG text output plus PNG band dimensions while retaining the richer-header/footer unsupported diagnostics.
146. Done shared image composition slice: add `OfficeImageComposer` and `OfficeImageLayer` in `OfficeIMO.Drawing` for dependency-free PNG/SVG page background, positioned layer placement, and optional raster/SVG adornment callbacks; migrate Excel print-title page assembly and header/footer chrome assembly onto it so page-level image composition has one shared implementation path.
147. Done Excel page-setup canvas slice: compose manual-page-sliced worksheet image results onto a shared Drawing-backed physical page canvas, applying orientation, margins, and manual scale, emitting `ExcelPageSetupPaperSizeDefaulted` for default Letter paper-size geometry, preserving explicit `ExcelPageSetupUnsupported` diagnostics for fit-to-width/fit-to-height, and proving PNG/SVG page dimensions plus margin placement through focused export tests.
148. Done shared physical page-size slice: add neutral `OfficePageSize` and `OfficePageSizes` primitives to `OfficeIMO.Drawing`, teach Excel page setup to read/write known OpenXML paper-size codes, route page-sliced image canvas dimensions through the shared physical-size primitive, diagnose unknown paper-size codes with `ExcelPageSetupPaperSizeUnsupported`, and prove A4 PNG, Legal landscape SVG, unknown-code fallback diagnostics, and print-layout paper-size persistence.
149. Done Excel page-sliced header/footer field slice: pass page number and page count through multi-image worksheet export, render plain header/footer text with supported `&P`, `&N`, `&A`, `&F`, `&Z`, `&&`, `&[Page]`, `&[Pages]`, `&[Tab]`, `&[File]`, and `&[Path]` fields in PNG/SVG page slices, and keep unsupported formatting/image fields source-diagnosed instead of silently approximated.
150. Done Excel page-sliced header/footer variant slice: select first-page, even-page, or odd/default plain text header/footer variants per page when `differentFirst` or `differentOddEven` is enabled, keep the same shared Drawing composition path for PNG/SVG output, and prove three-page first/even/odd SVG output with unsupported diagnostics absent.
151. Done Excel page-sliced header/footer date/time field slice: add deterministic `HeaderFooterDateTime` options for worksheet and workbook image export, render `&D`, `&T`, `&[Date]`, and `&[Time]` fields with culture-aware short date/time text in PNG/SVG page slices, and prove SVG output without unsupported diagnostics.
152. Done Excel page-sliced header/footer zone clipping slice: render plain header/footer left, center, and right sections inside non-overlapping zones, ellipsize overlong section text with shared text measurement/trimming, clip PNG text through shared `OfficeRasterCanvas` clip scopes, and clip SVG text through shared rectangular clip paths so narrow page slices do not let one section paint across another.
153. Done shared Drawing text-zone slice: move reusable three-column left/center/right zone layout and leading-ellipsis single-line trimming into `OfficeIMO.Drawing`; route Excel page-sliced header/footer PNG/SVG text zones through those shared contracts while keeping Excel-owned page fields, variant selection, clipping IDs, and unsupported-header/footer diagnostics in the adapter.
154. Done Excel header/footer basic formatting slice: parse `&B`, `&I`, and `&U` header/footer formatting tokens into styled runs, render them in PNG/SVG through shared Drawing rich-text measurement/segment emission, emit stable `ExcelHeaderFooterFormattingApproximation` diagnostics, and keep richer color/font/size/strikethrough formatting unsupported instead of silently flattening it.
155. Done shared rich-text strikethrough slice: add strikethrough to shared Drawing font/style, rich-run, rich-segment, raster text-line, and SVG text-decoration paths; route Excel header/footer `&S` through the shared rich-text path; keep header/footer font-family formatting and images diagnosed when unsupported; and prove Drawing text-decoration output plus Excel header/footer SVG output.
156. Done shared Drawing geometry rotation slice: move reusable degree/radian conversion and raster-space point rotation into `OfficeGeometry`; route `OfficeTextPlacement`, `OfficeRasterCanvas`, Visio PNG rotated text backgrounds, stencil artwork, and Visio raster adapter angle conversion through the shared primitive; tighten the Visio SVG cylinder test to assert curve semantics rather than path whitespace; and refresh/prove native Visio premium PNG/SVG baselines through the shared Drawing-backed gate.
157. Done Excel header/footer font-family slice: parse supported `&"Font[,Style]"` header/footer tokens into shared rich text runs, render supported SVG font-family/bold/italic output through `OfficeTextBlockRenderer.AppendSvgRichTextSegment`, keep malformed or richer style variants source-diagnosed instead of silently flattening them, and prove the supported and malformed paths with focused header/footer tests plus a reviewed PNG/SVG artifact.
158. Done shared raster rich-text renderer slice: add `OfficeTextBlockRenderer.DrawRasterRichTextBlock` for measured rich-run placement and raster dispatch, replace Excel cell rich-text and page-sliced header/footer PNG private segment loops with the shared Drawing helper, and prove Drawing run-color/style output plus Excel header/footer behavior.
159. Done shared raster font-family slice: add dependency-free CSS/Office font-family fallback resolution to `OfficeTrueTypeFont`, make `OfficeRasterCanvas` measure and draw with resolved family files when available, route Drawing scene text, Excel plain cell text, Excel rich cell text, and page-sliced header/footer text through the shared family-aware measurement/rendering path, and prove run font-family measurement plus guarded common-font fallback behavior.
160. Done Excel font diagnostic consolidation slice: use Drawing's stable `IMAGE_FONT_SUBSTITUTED` warning for explicit cell, rich-text, and header/footer families; resolve caller-scoped faces before platform fallback; and preserve Excel source references.
161. Done Excel chart font diagnostic consolidation slice: apply caller-scoped faces to the shared chart drawing before raster/SVG emission, diagnose concrete rendered chart text through Drawing, and remove the Excel-only chart font fallback code.
162. Done shared Word SVG image-detection slice: route Word file-backed SVG detection for the `a14:svgBlip` extension through `OfficeImageReader.FromExtension` from `OfficeIMO.Drawing` while keeping WordprocessingML picture wiring in Word; and prove saved DOCX SVG image parts still carry `image/svg+xml` and emit Office 2010 SVG blip markup.
163. Done shared image default-extension slice: add canonical image file-extension policy to `OfficeImageInfo`; route PowerPoint image part extension defaults through OpenXML part MIME -> Drawing format -> Drawing extension while preserving source-path extensions; and prove Drawing extension policy plus PowerPoint package-extension behavior.
164. Done shared image-extension recognition slice: add `OfficeImageReader.IsKnownImageExtension` as the shared dependency-free image-extension predicate; route Visio stencil preview relationship detection through it instead of a private extension switch; and prove Drawing-level extension recognition plus Visio external preview metadata discovered by target extension alone.
165. Done shared PDF image declared-type parity slice: add `OfficeImagePdfCompatibility.TryValidateDeclaredContentType` as the Drawing-owned policy for PDF-supported MIME type versus detected image bytes; route Excel PDF worksheet/header/footer image validation through it while leaving exact PNG stream encodability in the PDF writer; and prove Drawing mismatch diagnostics plus Excel PDF warning/skip behavior for declared-JPEG PNG bytes.
166. Done Excel fit-to-page image scaling slice: treat `FitToWidth` and `FitToHeight` values of `0` or `1` as bounded one-page content scaling inside the printable physical page canvas for manual-page-sliced PNG/SVG output, keep values above one page source-diagnosed as unsupported automatic multi-page fit pagination, and prove fit-to-width, fit-to-height, and remaining unsupported-diagnostic behavior with focused PNG/SVG export tests.
167. Done shared Excel page-geometry consolidation slice: add point conversion to `OfficePageSize`, move Excel paper-size resolution and fit-scale helpers into `ExcelPageSetupGeometry`, route page-sliced PNG/SVG export and first-party Excel PDF through that shared worksheet page setup resolver, preserve explicit PDF page-size precedence, and prove Drawing point conversion plus Excel image and PDF worksheet paper-size behavior with focused tests.
168. Done PDF page-background image-placement slice: route PDF page-background image fit rectangles through `OfficeImageRenderPlan.CreateBottomLeft` so flow, table, header/footer, canvas, and page-background image draw rectangles share the same Drawing-owned placement brain while PDF keeps page ordering, opacity graphics states, and XObject emission; also centralize PDF-owned image link-annotation bounds shared by flow and table-cell image rendering.
169. Done shared rotated scene-text slice: add rotation metadata and rotation-center tracking to `OfficeDrawingText`, route Drawing SVG scene text and rotated Drawing raster scene text through the shared text-block renderer, make Drawing visual-quality bounds understand rotated text boxes, and pass Excel worksheet drawing-object rotation into the shared Drawing text path. Excel keeps `ExcelDrawingShapeTextRotationApproximation` because text-box metrics are not Excel-exact yet, while the visible PNG/SVG output now rotates the shape label through the central Drawing engine and a dedicated approved visual baseline gates the result.
170. Done shared aligned scene-text slice: add vertical alignment metadata to `OfficeDrawingText`, route non-default Drawing scene-text placement through the shared text-block renderer for PNG/SVG, extract DrawingML paragraph alignment and body anchoring from Excel worksheet shapes into the neutral visual snapshot, and pass that metadata into the shared Drawing text path. A focused object test proves right/bottom DrawingML text alignment reaches PNG/SVG export, Drawing-level tests pin the shared aligned scene-text renderer, and a dedicated approved aligned shape-text PNG/SVG baseline was manually reviewed.
171. Done Excel DrawingML shape-text insets slice: extract `bodyPr` text wrapping plus left/top/right/bottom text insets with DrawingML defaults into the neutral worksheet drawing-object snapshot, pass the authored/default insets into the shared `OfficeDrawingText` scene-text rectangle instead of using an Excel-local padding heuristic, and prove default and authored zero-inset PNG/SVG export contracts on `net8.0` and `net472`.
172. Done shared scene-text shrink-to-fit slice: add opt-in shrink-to-fit metadata to `OfficeDrawingText`, route wrapped and non-wrapped Drawing scene text through existing shared fit-to-bounds layout primitives in PNG/SVG, extract Excel DrawingML `a:normAutofit` into the worksheet drawing-object snapshot, and prove Drawing-level SVG/PNG shrink plus public Excel shape-text PNG/SVG export contracts on `net8.0` and `net472`.
173. Done Excel DrawingML shape-text autofit diagnostic slice: extract `a:spAutoFit` into the neutral worksheet drawing-object snapshot as `TextResizeShapeToFit`, keep rendering through the shared fixed-bound Drawing scene-text path, and emit stable PNG/SVG diagnostics when authored shape-resize text behavior cannot be honored exactly.
174. Done Excel DrawingML shape-text vertical-orientation slice: map `bodyPr vert` values into a neutral `ExcelDrawingTextOrientation` snapshot enum, add opt-in stacked text intent to shared `OfficeDrawingText`, route simple DrawingML vertical shape text through shared stacked-text layout for PNG/SVG, gate the public range output with a dedicated approved vertical shape-text PNG/SVG baseline, and keep complex vertical variants such as `vert270`, East Asian, Mongolian, and WordArt orientations on stable PNG/SVG diagnostics until their semantics are implemented deliberately.
175. Done shared SVG rich-text whitespace slice: make `OfficeTextBlockRenderer` emit `xml:space="preserve"` when SVG text or rich-text segments contain leading, trailing, or repeated whitespace so run-boundary spaces survive in browser/viewer rendering; prove the Drawing-level segment contract plus Excel rich-text SVG export; refresh the affected approved rich-text, text-spill, and premium-range SVG baselines through the visual gate.
176. Done header/footer image visual QA slice: preserve multiple worksheet header/footer VML image positions when `SetHeaderImage` and `SetFooterImage` are called together, render the resulting PNG images through the existing shared Drawing page-composition path, and add a dedicated approved PNG/SVG visual baseline that gates the public page-sliced worksheet export output.
177. Done Excel visual-gate tiering slice: keep `Build/Test-ExcelImageVisualGate.ps1` defaulted to the full approved PNG/SVG baseline gate, add an explicit smoke suite for representative premium range, rich-text, header/footer image, chart, page-layout, conditional-formatting, drawing-object, and transformed-image baselines, and keep the shared Drawing architecture guard runnable as its own suite.
178. Done direct Drawing consumer guard slice: make `OfficeIMO.Excel.Pdf`, `OfficeIMO.Word.Pdf`, `OfficeIMO.Rtf.Pdf`, and `OfficeIMO.Markup.Word` reference `OfficeIMO.Drawing` directly instead of relying on transitive references, and add an architecture guard that fails any production `OfficeIMO.*` project that directly uses Drawing types without declaring the shared Drawing dependency explicitly.
179. Remaining depth: make visual fidelity production-grade: Excel-exact clipped, rotated, and stacked rich text; broader baseline font matrices and closer font metric parity; Excel-exact rotated/stacked text semantics; full conditional-formatting parity beyond bounded numeric comparisons including icon sets, date/time, and standard-deviation average rules; richer differential formats; exact complex/path/multi-stop gradient fills, more exact pattern fill density/parity, richer custom/scientific/conditional number-format display parity, richer image decoding/effects; exact image two-cell anchor clipping, transformations/effects beyond basic crop/flip/rotation, and hidden-row/column behavior as rendered parity beyond the current explicit diagnostics; more exact chart fidelity including picture markers and richer marker outline effects, custom/richer series dash and effect styling, richer point-level overrides, axis/tick formatting beyond simple value-axis numeric formats and simple high/low/next-to/none tick-label placement, Excel-exact display-unit placement/typography, remaining Excel-exact minor-gridline/tick placement edge cases, custom dash/effect parity beyond preset gridline and axis lines, chart/plot area effects beyond simple solid RGB fill/outline/width/preset dash, richer chart title typography/effects beyond simple font-family/font-size/bold/italic, per-element chart text runs beyond the supported shared buckets, and Excel-exact geometry; richer page slicing beyond manual breaks, repeated titles, basic header/footer text fields and approximate header/footer PNG images, bounded one-page fit scaling, and shared page geometry; automatic multi-page fit pagination; Excel-exact comment/note/threaded-comment popover geometry, threading, author metadata, visibility state, connectors, and stacking beyond the current opt-in callout approximation; Excel-exact sparkline parity for hidden/empty data and date-axis/group scaling behavior beyond the current shared renderer; richer shapes/text boxes/connectors including exact shape text autofit and complex vertical text orientations beyond simple stacked `vert`; deeper grouped-object/layer clipping/baseline metrics; broader visual-baseline matrices beyond the current rich-text and pattern-fill fixtures; and continue reusing Excel PDF planning internals where it reduces duplication.

## Consolidation Goal: One Rendering Brain

The larger goal is to move, replace, or delete duplicate rendering code until OfficeIMO has one dependency-free shared rendering brain:

```text
Office document package
  -> document-specific visual snapshot and source semantics
  -> OfficeIMO.Drawing primitives, raster canvas, SVG/PNG exporters, diagnostics
  -> document-specific API convenience wrappers
```

This means:

- `OfficeIMO.Drawing` owns pixels, paths, fills, strokes, clipping, text layout primitives, image decode/encode where dependency-free support exists, chart/drawing primitives, reusable diagnostics, and visual-quality helpers.
- `OfficeIMO.Excel` owns workbook/worksheet/range semantics, Open XML extraction, Excel-specific layout policy, Excel number/date/style interpretation, and friendly Excel image APIs.
- `OfficeIMO.Visio` owns VSDX page, shape, connector, stencil, routing, and coordinate semantics, but not reusable PNG encoding, raster buffers, generic stroke/path/text/image projection loops, or duplicate text-layout math.
- `OfficeIMO.Pdf`, `OfficeIMO.Word`, and `OfficeIMO.PowerPoint` may have format-specific writers and layout semantics, but should reuse Drawing primitives whenever they need image-like rendering.

The anti-pattern is a second product renderer that quietly grows a private answer for clipping, dashed strokes, image transforms, text wrapping, or diagnostics. When Excel needs something Visio already solved, move the reusable part to `OfficeIMO.Drawing`, keep the source-format policy in the document adapter, and prove at least one non-Excel consumer still works.

### Migration Order

1. Inventory current renderers and classify each piece as shared engine, document adapter, PDF-specific writer behavior, test helper, or duplicate private renderer.
2. Finish retiring Visio private raster primitives into `OfficeIMO.Drawing`, including any remaining reusable path, clipping, text, image, or stroke behavior that is not inherently Visio-specific.
3. Keep Visio premium baselines passing while migration happens, because Visio is the strongest current proof that shared Drawing can render polished visuals.
4. Keep Excel range/worksheet/workbook export using the same Drawing engine and shared diagnostics; do not add Excel-only pixel/path/text engines.
5. Consolidate test helpers so Excel, Visio, PDF, and later PowerPoint/Word visual baselines compare through one approved-image workflow.
6. Add new Excel premium features only after deciding whether their generic part belongs in Drawing and their source-policy part belongs in Excel.
7. When a feature cannot be rendered with parity, add a stable diagnostic code with source reference before improving visual output.
8. Delete or shrink the old private implementation after its shared replacement is proven by contract tests and visual baselines.

Current consolidation checkpoint: Visio SVG and PNG text rendering share `OfficeTextBlockRenderPlan` for fitted center-based text boxes, aligned placement, text background bounds, and rotated background corners; Excel PNG/SVG plain cell text shares the same Drawing-owned render plan for left/top cell rectangles, wrapped/stacked layout, shrink-to-fit sizing, and vertical placement; and Excel rich-text SVG output now uses Drawing-owned rich-text block emission beside the existing shared raster rich-text renderer. Visio still owns VSDX coordinate/style semantics and Excel still owns workbook style/rotation/diagnostic semantics, while `OfficeIMO.Drawing` owns the reusable text-block placement and rich-run emission math both adapters need.

## Premium Excel Export Goal

Premium Excel export means PNG/SVG output that looks intentionally rendered, not hand-drawn, while still being honest about the places where Office-exact parity is not implemented. The goal is not byte-identical desktop Excel screenshots. The goal is deterministic, dependency-free OfficeIMO rendering with strong fidelity, stable diagnostics, and visual QA gates.

### Premium Workstreams

- Text/layout: Excel-like wrapping, clipping, vertical alignment, shrink-to-fit, rotated and stacked text, rich text runs, baseline metrics, font fallback, and measured-text caching belong in shared Drawing where possible; Excel keeps style and layout policy.
- Styles: theme colors, tints/shades, number/date/currency/percent display text including custom literal/section formats, conditional formatting, border styles, gradients, and pattern fills should be source-resolved by Excel and painted by Drawing.
- Images: embedded image metadata, formats, aspect behavior, anchors, clipping, crop, transparency, transforms, z-order, and SVG embedding/raster decoding should flow through one Drawing image path with Excel source diagnostics.
- Charts: Excel chart extraction maps authored style/layout into shared chart snapshots; Drawing renders the chart primitives; diagnostics explain trendlines, leader lines, rich chart text, exact axis behavior, and effects that are not yet parity-grade.
- Worksheet/page behavior: ranges, used ranges, print areas, page setup, orientation, scaling, page slicing, large-sheet tiling, headers/footers where in scope, and hidden rows/columns are Excel semantics over the same renderer.
- Objects: shapes, text boxes, connectors, comments/notes, threaded comments, hyperlinks, sparklines, and drawing layers enter one ordered visual snapshot and one renderer path.
- Diagnostics: every unsupported or approximated feature needs a stable code, severity, human-readable message, and source reference where the workbook provides one.
- Visual QA: premium work needs approved baselines, focused decoded-pixel/SVG assertions, and manual review of saved artifacts when the output changes meaningfully.

### Premium Done Criteria

- Supported output is visually credible at normal report/screenshot scale and does not look like debug rendering.
- PNG and SVG are visually comparable for the same source within the limits of each format.
- Source ordering, clipping, and z-order are explicit contracts for cells, shapes, images, and charts.
- Unsupported features never silently disappear.
- Each premium slice either improves shared Drawing or proves why the feature is truly Excel-specific.
- Baseline failures produce reviewable diff artifacts, not just red tests.
- Large ranges and worksheets have a documented memory/page/tiling story before workbook-wide export is called production-grade.

### Continuation Checkpoint 2026-06-25

This PR proves the right architecture and a credible first Excel PNG/SVG export path, but it should not be described as premium-complete. The shared renderer now has approved visual baselines, a smoke/full visual gate, stable diagnostics, and architecture guards that keep Excel, Visio, and PDF-oriented rendering paths using `OfficeIMO.Drawing` instead of separate private brains. The remaining work is a fidelity burn-down, not a rewrite.

Current approved Excel image baselines are split into clean baselines and tracked approximations by `ExcelImageExportVisualFidelityGateTests`. The tracked approximation set is the authoritative continuation list for the next premium PRs:

- Comment and threaded-comment bodies: render as dependency-free callouts, but still need Excel-exact popover geometry, visibility state, threading metadata, stacking, connector/pointer behavior, and richer body formatting.
- Conditional formatting icons: deterministic and visually improved, but still need Excel icon-artwork parity, threshold semantics, richer icon sets, and diagnostics for any impossible parity.
- Pattern fills and gradients: hatch output is deterministic, but needs closer Excel pattern density/phase/color behavior, broader gradient handling, and theme-aware parity.
- Rich text and cell text layout: readable output exists, including wrapping, clipping, shrink-to-fit, rotated text, and stacked text, but premium still needs Excel-exact baseline metrics, font fallback behavior, vertical alignment edge cases, text spill/overflow parity, and a measured-text cache story for larger exports.
- Drawing-object text: supported simple shapes route through Drawing, but rotated shape text, text-box metrics, auto-fit, complex vertical orientation, connectors, groups, theme/system/transformed colors, and richer preset geometry remain incomplete.
- Sparklines: line, column, and win/loss output exists, but hidden/empty data behavior, date axes, group-level scaling, axis display, and Excel-exact marker/negative/first/last/high/low semantics need deeper parity.
- Charts: the current snapshot bridge carries a useful slice of authored style/layout into shared Drawing, but premium still needs trendlines, leader lines, point-level label overrides, picture markers, richer series/marker effects, deeper axis/tick/number-format behavior, display-unit placement, chart/plot area effects, per-element rich chart text, and Excel-exact chart geometry where practical.
- Images and objects: embedded PNG/SVG-friendly images, crop, rotation, flip, two-cell sizing, clipping, and z-order are covered, but broader raster formats, EMF/SVG/JPEG edge cases, transparency/effects, grouped objects, connectors, and deeper object clipping still need explicit contracts.
- Worksheet/page export: print area, manual page slicing, print titles, page setup, orientation, margins, scaling, paper sizes, and text header/footer chrome have first-pass support, but automatic pagination, large-sheet tiling, broader paper-size coverage, Excel-exact header/footer image geometry, and full page-break/fit-to-page parity are still open.
- Diagnostics and QA: every unsupported or approximate detail must keep a stable diagnostic code with a source reference. Visual QA must expand beyond curated baselines into broader fixture matrices, renderable/nonblank checks, visual diff artifacts, and, where practical, Excel-reference comparison notes.

Recommended next PR order:

1. Keep the multi-target build lane clean, especially `net472`, before starting new fidelity work.
2. Burn down the most visible tracked approximations first: comment bodies, rotated/stacked text metrics, sparklines, pattern fills, and conditional icons.
3. Expand chart parity in focused slices, with diagnostics for anything deliberately approximate.
4. Add larger worksheet/page tiling and pagination contracts only after range and page-sliced output stay stable.
5. Continue moving reusable geometry, text, image projection, and SVG/raster primitives into `OfficeIMO.Drawing`; keep Excel, Visio, PDF, Word, and PowerPoint as source-format adapters over the shared engine.

The first visible milestone can be:

```csharp
sheet.Range("A1:D12").SaveAsPng("range.png");
sheet.Range("A1:D12").SaveAsSvg("range.svg");
```

That is intentionally small, but the internal path should already be:

```text
Excel package -> Excel visual snapshot -> shared renderer -> PNG/SVG encoder
```

## Definition Of Done For Each Phase

- No runtime dependencies added.
- Existing package behavior remains intact.
- Public APIs are documented with realistic examples.
- Diagnostics describe unsupported or approximate rendering.
- Tests prove valid PNG/SVG output, stable dimensions, and nonblank visual content.
- Visual review proves exported artifacts look professionally rendered and not hand-drawn/debug-like.
- Renderer changes include either approved visual baselines or saved QA artifacts that exercise cells, text, borders, fills, images, and charts.
- When a phase touches shared rendering, at least one non-Excel consumer or fixture proves the code is not Excel-specific.

## Current Supporting Assessment

The detailed current-state assessment is in `Docs/reviews/officeimo.excel-image-conversion-assessment-2026-06-22.md`.
