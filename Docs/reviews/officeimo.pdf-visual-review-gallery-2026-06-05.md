# OfficeIMO PDF Visual Review Gallery - 2026-06-05

## Purpose

This gallery is the manual review companion for the Poppler raster baseline suite. It points reviewers to actual generated PDFs, not screenshots, and explains what each artifact is meant to prove visually.

Generate the current artifact pack with:

```powershell
Build/Export-PdfVisualReviewGallery.ps1 -OutputDirectory artifacts/pdf-visual-review -Framework net8.0
```

Add `-RequireRasterizer` when the run must fail if Poppler `pdftoppm` is not
available. The script writes PDFs plus an `index.md` file that records the
commit, output directory, command, and generated file list.

The original manual review artifacts for this snapshot lived under:

`C:\Support\GitHub\OfficeIMO-pdf-implementation-review-20260605\artifacts\pdf-visual-review`

## Review Checklist

- Open the PDFs in Edge and Acrobat/Reader when possible.
- Check text smoothness at several zoom levels, especially 90%, 100%, 125%, and fit-to-width.
- Check that images preserve aspect ratio unless the scenario intentionally uses stretch.
- Check that charts, tables, and lists stay inside their frames and have comfortable spacing.
- Check that TOC leaders, links, and table borders do not cross columns or neighboring content.
- Check that branded/header/footer images are not distorted and do not overlap text.
- Check that colors retain contrast on white and tinted backgrounds.

## Gallery

| File | Scenario | What to Inspect |
| --- | --- | --- |
| `01-professional-report.pdf` | Report authoring | Report rhythm, headings, wrapped table cells, image placement, bookmarks. |
| `02-showcase-dashboard.pdf` | Dashboard | Dense cards, charts, numeric hierarchy, color contrast, repeated layout rhythm. |
| `03-line-items-two-page.pdf` | Statement/invoice table | Multi-page table flow, row spacing, borders, footer continuity. |
| `04-flow-dsl.pdf` | Core compose DSL | Embedded Arial rendering, colored inline text, multi-page flow, footer placement. |
| `05-lists-tables.pdf` | Lists and tables | List indentation, marker placement, table padding, borders, wrapped cells. |
| `06-table-style-gallery.pdf` | Table styles | Header/body contrast, row striping, border consistency, compact spacing. |
| `07-drawing-gallery.pdf` | Shared vector drawing | Embedded Arial rendering, gradients, shadows, dashed strokes, clipping, transforms. |
| `08-row-columns.pdf` | Core row/column layout | Column widths, gutters, alignment, paragraph wrapping inside columns. |
| `09-background-image.pdf` | Page background image | Image fit, opacity, text contrast over image-backed pages. |
| `10-background-shapes.pdf` | Page background shapes | Decorative bands/shapes, z-order, margins, non-overlap with content. |
| `11-native-word-report.pdf` | Word conversion report | Word headings, table mapping, lists, links, image flow, native export fidelity. |
| `12-native-word-daily-layout.pdf` | Word daily layout gate | Embedded Calibri, TOC leader bounds, two-column flow, table placement, logo sizing. |
| `13-native-word-table-cell-picture-control.pdf` | Word table images | Table-cell image sizing, text wrapping beside images, cell padding. |
| `14-native-excel-daily-workbook.pdf` | Excel workbook conversion | Sheet pagination, chart placement, table chunking, images, hyperlinks. |
| `15-markdown-technical-document.pdf` | Markdown technical doc | Code blocks, headings, links, table/list rhythm, long text wrapping. |
| `16-markdown-theme-report.pdf` | Markdown themed report | Theme colors, spacing, headings, panels, table readability. |
| `17-native-powerpoint-slide.pdf` | PowerPoint slide conversion | Slide size mapping, proportional picture fit, text boxes, shapes, tables, chart placement. |
| `18-native-powerpoint-dense-layout.pdf` | Dense PowerPoint conversion | Clipping diagnostics, dense chart labels, list/table spacing, warnings coverage. |

## Current Phase 3/4 Notes

- Core text rendering now clamps aligned tab leaders to the active frame, preventing Word TOC leaders from crossing columns.
- Word-origin PDFs embed the document default font when available, stabilizing viewer rendering and metrics.
- PowerPoint slide background images use proportional cover fitting instead of stretching.
- Uncropped PowerPoint pictures default to proportional contain fitting. Explicit stretch remains available through `PowerPointPdfSaveOptions.PictureFit`.
- PowerPoint export records a `picture-aspect-distortion` warning when explicit stretch is used in a frame with a visibly different aspect ratio.

## Validation Commands

```powershell
dotnet test OfficeIMO.Tests\OfficeIMO.Tests.csproj --framework net8.0 --filter "FullyQualifiedName~PdfDocumentRasterVisualBaselineTests" --no-restore --verbosity minimal -p:WarningLevel=0 -clp:ErrorsOnly
dotnet test OfficeIMO.Tests\OfficeIMO.Tests.csproj --framework net8.0 --filter "FullyQualifiedName~OfficeIMO.Tests.Pdf.PowerPointSaveAsPdfTests" --no-restore --verbosity minimal -p:WarningLevel=0 -clp:ErrorsOnly
```
