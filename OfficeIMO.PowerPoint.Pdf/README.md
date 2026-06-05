# OfficeIMO.PowerPoint.Pdf

First-party PowerPoint-to-PDF export built on the reusable `OfficeIMO.Pdf` engine.

Current scope:

- one PowerPoint slide maps to one PDF page using the slide size in points
- ordered slide canvas rendering for slide backgrounds, text boxes, supported pictures, supported tables, supported charts, and basic auto-shapes
- solid, cover-fit image, and two-stop linear gradient slide backgrounds reuse the PowerPoint background snapshot API plus shared PDF shape/image primitives
- text boxes map fill, outline, margins, font defaults, horizontal alignment, vertical anchoring, rich runs, and hyperlinks to `PdfCanvasTextBoxStyle`
- pictures reuse the shared PDF JPEG/PNG image pipeline; uncropped pictures default to proportional `Contain` fitting through `PowerPointPdfSaveOptions.PictureFit`, while cropped pictures preserve authored crop-frame semantics
- tables reuse the shared fixed-position PDF canvas table primitive with authored column widths, row heights, fills, padding, alignment, simple text styling, merge spans, and basic borders
- clustered column, line, scatter, pie, and doughnut charts use the PowerPoint chart snapshot API plus the shared `OfficeChartDrawingRenderer` vector chart renderer, with optional `PowerPointPdfSaveOptions.ChartStyle` and `ChartLayout` pass-through for consistent chart palette, text, and dense-label layout
- rectangle, rounded rectangle, ellipse, and line auto-shapes reuse `OfficeIMO.Drawing` descriptors
- unsupported slide content, stretched-picture aspect ratio diagnostics, and shared chart drawing quality diagnostics record warnings through `PowerPointPdfSaveOptions.Warnings`

This package is intentionally a thin adapter over the shared PDF engine. Richer slide fidelity should land in `OfficeIMO.Pdf` reusable primitives first, then be exposed here.
