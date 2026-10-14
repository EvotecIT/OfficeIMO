using System;
using System.Collections.Generic;
using System.Linq;
using OfficeIMO.Drawing;
using PdfCore = OfficeIMO.Pdf;
using PptCore = OfficeIMO.PowerPoint;

namespace OfficeIMO.PowerPoint.Pdf;

public static partial class PowerPointPdfConverterExtensions {
    private static void RenderShapes(PdfCore.PdfPageCanvas canvas, IReadOnlyList<PptCore.PowerPointShape> shapes, int slideNumber, double pageWidth, double pageHeight, PowerPointToPdfOptions options, bool warnInvalidBounds, int groupDepth) {
        foreach (PptCore.PowerPointShape shape in shapes) {
            if (shape.Hidden) {
                continue;
            }

            if (!TryGetShapeBox(shape, slideNumber, pageWidth, pageHeight, options, warnInvalidBounds, out double x, out double y, out double width, out double height)) {
                continue;
            }

            Action<PdfCore.PdfPageCanvas> render = target => RenderShapeContent(target, shape, x, y, width, height, slideNumber, pageWidth, pageHeight, options, warnInvalidBounds, groupDepth);
            if (shape is PptCore.PowerPointGroupShape || IsLineShape(shape)) {
                // Establish the page clip before nested transforms or a line's
                // stroke-sized link region extend beyond the raw shape box.
                canvas.Clip(0D, 0D, pageWidth, pageHeight, render);
            } else if (TryGetVisibleSlideBox(x, y, width, height, pageWidth, pageHeight, out double clipX, out double clipY, out double clipWidth, out double clipHeight) &&
                NeedsSlideClip(x, y, width, height, pageWidth, pageHeight)) {
                canvas.Clip(clipX, clipY, clipWidth, clipHeight, render);
            } else {
                render(canvas);
            }
        }
    }

    private static void RenderShapeContent(PdfCore.PdfPageCanvas canvas, PptCore.PowerPointShape shape, double x, double y, double width, double height, int slideNumber, double pageWidth, double pageHeight, PowerPointToPdfOptions options, bool warnInvalidBounds, int groupDepth) {
        if (shape is not PptCore.PowerPointPicture && shape.Hyperlink is { IsAbsoluteUri: true })
            AddWarning(options, slideNumber, "shape-hyperlink-omitted", "The PowerPoint shape hyperlink was not exported as a PDF annotation.", lossKind: OfficeConversionLossKind.Omission);
        if (shape is PptCore.PowerPointTextBox textBox) {
            bool renderedGeometry = options.IncludeAutoShapes && RenderTextBoxGeometry(canvas, textBox, x, y, width, height);
            if (options.IncludeTextBoxes) {
                RenderTextBox(canvas, textBox, x, y, width, height, slideNumber, options, suppressFrame: renderedGeometry);
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
                RenderTable(canvas, table, x, y, width, height, slideNumber, options);
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
                RenderSmartArt(canvas, smartArt, x, y, width, height,
                    slideNumber, options);
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

        AddWarning(options, slideNumber, "unsupported-shape", "Skipped unsupported PowerPoint shape content type '" + shape.ShapeContentType + "'.", lossKind: OfficeConversionLossKind.Omission);
    }

}
