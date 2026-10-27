using System;
using System.Collections.Generic;
using System.Linq;
using OfficeIMO.Drawing;
using PdfCore = OfficeIMO.Pdf;
using PptCore = OfficeIMO.PowerPoint;

namespace OfficeIMO.PowerPoint.Pdf;

public static partial class PowerPointPdfConverterExtensions {
    private static void RenderShapes(PdfCore.PdfPageCanvas canvas, IReadOnlyList<PptCore.PowerPointShape> shapes, int slideNumber, double pageWidth, double pageHeight, PowerPointPdfExportOptions options, bool warnInvalidBounds, int groupDepth) {
        foreach (PptCore.PowerPointShape shape in shapes) {
            if (shape.Hidden && !options.IncludeHiddenShapes) {
                continue;
            }

            if (!TryGetShapeBox(shape, slideNumber, pageWidth, pageHeight, options, warnInvalidBounds, out double x, out double y, out double width, out double height)) {
                continue;
            }

            Action<PdfCore.PdfPageCanvas> render = target => RenderShapeContent(target, shape, x, y, width, height, slideNumber, pageWidth, pageHeight, options, warnInvalidBounds, groupDepth);
            // Clip against the slide, never the original shape box: autofit rows
            // and fixed-frame text may legitimately paint beyond that box.
            if (shape is PptCore.PowerPointGroupShape || IsLineShape(shape) ||
                x < 0D || y < 0D || x + width > pageWidth || y + height > pageHeight)
                canvas.Clip(0D, 0D, pageWidth, pageHeight, render);
            else render(canvas);
        }
    }

    private static void RenderShapeContent(PdfCore.PdfPageCanvas canvas, PptCore.PowerPointShape shape, double x, double y, double width, double height, int slideNumber, double pageWidth, double pageHeight, PowerPointPdfExportOptions options, bool warnInvalidBounds, int groupDepth) {
        if (shape is not PptCore.PowerPointPicture && shape.Hyperlink is { IsAbsoluteUri: true })
            AddWarning(options, slideNumber, "shape-hyperlink-omitted", "The PowerPoint shape hyperlink was not exported as a PDF annotation.", lossKind: OfficeConversionLossKind.Omission);
        if (shape is PptCore.PowerPointTextBox textBox) {
            bool renderedGeometry = false;
            if (options.IncludeTextBoxes) {
                RenderTextBox(canvas, textBox, x, y, width, height, slideNumber, options);
            } else if (options.IncludeAutoShapes) {
                renderedGeometry = RenderTextBoxGeometry(canvas, textBox, x, y, width, height, slideNumber, options);
            }
            if (renderedGeometry || options.IncludeTextBoxes)
                RenderShapeSlideLink(canvas, shape, x, y, width, height, slideNumber, options);
            return;
        }

        if (shape is PptCore.PowerPointPicture picture) {
            if (options.IncludePictures) {
                RenderPicture(canvas, picture, x, y, width, height, slideNumber, options);
            }
            return;
        }

        if (shape is PptCore.PowerPointTable table) {
            if (options.IncludeTables) {
                // Group children retain authored coordinates under their affine;
                // compare against page bounds only in the slide coordinate space.
                RenderTable(canvas, table, x, y, width, height, slideNumber, options,
                    groupDepth == 0 ? pageWidth : null, groupDepth == 0 ? pageHeight : null, hasAncestorPageClip: groupDepth > 0);
                RenderShapeSlideLink(canvas, shape, x, y, width, height, slideNumber, options);
            }
            return;
        }

        if (shape is PptCore.PowerPointChart chart) {
            if (options.IncludeCharts) {
                RenderChart(canvas, chart, x, y, width, height, slideNumber, options);
                RenderShapeSlideLink(canvas, shape, x, y, width, height, slideNumber, options);
            }
            return;
        }

        if (shape is PptCore.PowerPointSmartArt smartArt) {
            if (options.IncludeSmartArt) {
                // Office caches the laid-out SmartArt shapes for exactly this purpose.
                if (smartArt.TryCreateCachedDrawing(out PptCore.PowerPointGroupShape cachedDrawing, out string? cachedDrawingFailure)) {
                    RenderGroupShape(canvas, cachedDrawing, x, y, width, height, slideNumber, pageWidth, pageHeight, options, warnInvalidBounds, groupDepth);
                } else {
                    RenderSmartArt(canvas, smartArt, x, y, width, height,
                        slideNumber, options, cachedDrawingFailure);
                }
                RenderShapeSlideLink(canvas, shape, x, y, width, height, slideNumber, options);
            }
            return;
        }

        if (shape is PptCore.PowerPointConnectionShape connection) {
            if (options.IncludeAutoShapes) {
                RenderConnection(canvas, connection, x, y, width, height, slideNumber, options);
                RenderShapeSlideLink(canvas, shape, x, y, width, height, slideNumber, options);
            }
            return;
        }

        if (shape is PptCore.PowerPointAutoShape autoShape) {
            if (options.IncludeAutoShapes) {
                RenderAutoShape(canvas, autoShape, x, y, width, height, slideNumber, options);
                RenderShapeSlideLink(canvas, shape, x, y, width, height, slideNumber, options);
            }
            return;
        }

        if (shape is PptCore.PowerPointGroupShape groupShape) {
            if (shape.OwnerSlide != null) {
                RenderGroupShape(canvas, groupShape, x, y, width, height, slideNumber, pageWidth, pageHeight, options, warnInvalidBounds, groupDepth);
                RenderShapeSlideLink(canvas, shape, x, y, width, height, slideNumber, options);
            } else {
                AddWarning(options, slideNumber, "unsupported-shape", "Skipped a PowerPoint group shape because its owning slide context could not be resolved.", lossKind: OfficeConversionLossKind.Omission);
            }
            return;
        }

        if (shape is PptCore.PowerPointOleObject) {
            AddWarning(options, slideNumber, "unsupported-shape", "Skipped a PowerPoint OLE object because an embedded-object preview is unavailable.",
                lossKind: OfficeConversionLossKind.Omission);
            return;
        }

        AddWarning(options, slideNumber, "unsupported-shape", "Skipped unsupported PowerPoint shape content type '" + shape.ShapeContentType + "'.", lossKind: OfficeConversionLossKind.Omission);
    }

}
