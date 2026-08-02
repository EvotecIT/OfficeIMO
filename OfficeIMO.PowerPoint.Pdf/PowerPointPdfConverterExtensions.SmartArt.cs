using System;
using OfficeIMO.Drawing;
using PdfCore = OfficeIMO.Pdf;
using PptCore = OfficeIMO.PowerPoint;

namespace OfficeIMO.PowerPoint.Pdf;

public static partial class PowerPointPdfConverterExtensions {
    private static void RenderSmartArt(PdfCore.PdfPageCanvas canvas,
        PptCore.PowerPointSmartArt smartArt, double x, double y,
        double width, double height, int slideNumber,
        PowerPointPdfSaveOptions options) {
        if (!smartArt.TryGetOfficeDiagramSnapshot(
                out OfficeDiagramSnapshot source)) {
            AddLayoutWarning(options, slideNumber, "unsupported-smartart",
                "Skipped a PowerPoint SmartArt diagram because its semantic node data could not be read safely.",
                PdfCore.PdfLayoutDiagnosticKind.SkippedContent,
                "PowerPointSmartArt",
                "The SmartArt semantic model could not be read into the shared PDF diagram renderer.",
                x, y, width, height);
            return;
        }

        try {
            var snapshot = new OfficeDiagramSnapshot(source.Name,
                source.Kind, source.Nodes, width, height, source.Style);
            OfficeDrawing drawing = OfficeDiagramDrawingRenderer.Render(
                snapshot, includeBackground: false);
            var frameTransform = new OfficeImageFrameTransform(
                smartArt.Rotation ?? 0D, width / 2D, height / 2D,
                smartArt.HorizontalFlip == true,
                smartArt.VerticalFlip == true);
            if (frameTransform.HasTransform) {
                var transformed = new OfficeDrawing(width, height);
                transformed.AddDrawing(drawing, 0D, 0D, frameTransform);
                drawing = transformed;
            }
            canvas.Drawing(drawing, x, y, width, height,
                style: new PdfCore.PdfDrawingStyle {
                    AlternativeText = string.IsNullOrWhiteSpace(source.Name)
                        ? "PowerPoint SmartArt diagram"
                        : source.Name
                });
        } catch (Exception exception) when (exception is ArgumentException
            or InvalidOperationException) {
            AddLayoutWarning(options, slideNumber, "unsupported-smartart",
                "Skipped a PowerPoint SmartArt diagram because it could not be rendered as a shared PDF diagram: "
                    + exception.Message,
                PdfCore.PdfLayoutDiagnosticKind.SkippedContent,
                "PowerPointSmartArt",
                "The SmartArt diagram could not be rendered by the shared PDF diagram renderer.",
                x, y, width, height);
        }
    }
}
