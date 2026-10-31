using System;
using System.Collections.Generic;
using System.Linq;
using OfficeIMO.Drawing;
using PdfCore = OfficeIMO.Pdf;
using PptCore = OfficeIMO.PowerPoint;

namespace OfficeIMO.PowerPoint.Pdf;

public static partial class PowerPointPdfConverterExtensions {
    private static void RenderTextBox(PdfCore.PdfPageCanvas canvas, PptCore.PowerPointTextBox textBox, double x, double y, double width, double height, int slideNumber, PowerPointPdfExportOptions options) {
        var diagnostics = new List<OfficeImageExportDiagnostic>();
        OfficeDrawing drawing = PptCore.PowerPointSlideImageRenderer.CreateTextBoxDrawing(
            textBox, width, height, diagnostics, options.NativeTextMeasure!, options.NativeTextDefaultFontFamily, options.CancellationToken);
        bool renderedGeometry = options.IncludeAutoShapes && RenderTextBoxGeometry(
            canvas, textBox, x, y, drawing.Width, drawing.Height, slideNumber, options);
        if (!renderedGeometry) RenderTextBoxFrame(canvas, textBox, x, y, drawing.Width, drawing.Height, slideNumber, options);
        foreach (OfficeImageExportDiagnostic diagnostic in diagnostics) {
            string code = diagnostic.Code switch {
                "powerpoint-text-overflow" => "text-box-overflow",
                "powerpoint-text-padding" => "text-box-padding",
                PptCore.PowerPointImageExportDiagnosticCodes.SmallCapsApproximated => "small-caps-approximation",
                PptCore.PowerPointImageExportDiagnosticCodes.BaselinePercentApproximated => "baseline-percent-approximation",
                _ => diagnostic.Code
            };
            if (code == "text-box-overflow") {
                AddWarning(options, slideNumber, code, diagnostic.Message,
                    new PdfCore.PdfLayoutDiagnostic(PdfCore.PdfLayoutDiagnosticKind.Overflow, "PowerPointTextBox",
                        diagnostic.Message, global::OfficeIMO.OfficeLength.Points(x), global::OfficeIMO.OfficeLength.Points(y), global::OfficeIMO.OfficeLength.Points(width), global::OfficeIMO.OfficeLength.Points(height)), diagnostic.LossKind);
            } else {
                AddWarning(options, slideNumber, code, diagnostic.Message, lossKind: diagnostic.LossKind);
            }
        }
        foreach (PptCore.PowerPointTextRun run in textBox.Paragraphs.SelectMany(paragraph => paragraph.InlineNodes)
                     .Where(node => node.Run != null).Select(node => node.Run!)) {
            if (run.Hyperlink is { IsAbsoluteUri: false } link)
                ResolveSlideLinkDestination(link, slideNumber, options, "text run");
        }
        RemoveUnavailableDrawingSlideLinks(drawing, options);
        // The authored frame is not a clipping boundary for PowerPoint text.
        // Retain finite overflow coordinates without enlarging or rescaling that frame.
        if (drawing.Elements.Count > 0) canvas.DrawingForClippedRendering(drawing, x, y, drawing.Width, drawing.Height);
    }

    private static void RenderTextBoxFrame(PdfCore.PdfPageCanvas canvas, PptCore.PowerPointTextBox textBox, double x, double y, double width, double height,
        int slideNumber, PowerPointPdfExportOptions options) {
        OfficeShape frame = OfficeShape.Rectangle(width, height);
        ApplyShapeAppearance(textBox, frame, slideNumber, options);
        if (!frame.FillColor.HasValue && frame.FillGradient == null && frame.FillRadialGradient == null
            && !frame.StrokeColor.HasValue) return;
        PptCore.PowerPointSlideImageRenderer.ApplyShapeTransformForExport(frame, textBox, width, height);
        canvas.Shape(frame, global::OfficeIMO.OfficeLength.Points(x), global::OfficeIMO.OfficeLength.Points(y));
    }

}
