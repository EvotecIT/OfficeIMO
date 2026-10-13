using System;
using System.Collections.Generic;
using System.Linq;
using OfficeIMO.Drawing;
using PdfCore = OfficeIMO.Pdf;
using PptCore = OfficeIMO.PowerPoint;

namespace OfficeIMO.PowerPoint.Pdf;

public static partial class PowerPointPdfConverterExtensions {
    private static void RenderTextBox(PdfCore.PdfPageCanvas canvas, PptCore.PowerPointTextBox textBox, double x, double y, double width, double height, int slideNumber, PowerPointToPdfOptions options, bool suppressFrame = false) {
        if (!suppressFrame) RenderTextBoxFrame(canvas, textBox, x, y, width, height);
        var diagnostics = new List<OfficeImageExportDiagnostic>();
        OfficeDrawing drawing = PptCore.PowerPointSlideImageRenderer.CreateTextBoxDrawing(
            textBox, width, height, diagnostics, options.NativeTextMeasure!, options.NativeTextDefaultFontFamily);
        foreach (OfficeImageExportDiagnostic diagnostic in diagnostics) {
            string code = diagnostic.Code switch {
                "POWERPOINT_TEXT_OVERFLOW" => "text-box-overflow",
                "POWERPOINT_TEXT_PADDING" => "text-box-padding",
                PptCore.PowerPointImageExportDiagnosticCodes.SmallCapsApproximated => "small-caps-approximation",
                PptCore.PowerPointImageExportDiagnosticCodes.BaselinePercentApproximated => "baseline-percent-approximation",
                _ => diagnostic.Code
            };
            if (code == "text-box-overflow") {
                AddLayoutWarning(options, slideNumber, code, diagnostic.Message,
                    PdfCore.PdfLayoutDiagnosticKind.ClippedContent, "PowerPointTextBox",
                    diagnostic.Message, x, y, width, height);
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
        if (drawing.Elements.Count > 0) canvas.Drawing(drawing, x, y, width, height);
    }

    private static void RenderTextBoxFrame(PdfCore.PdfPageCanvas canvas, PptCore.PowerPointTextBox textBox, double x, double y, double width, double height) {
        OfficeColor? fill = textBox.FillTransparency == 100 ? null : ParseOfficeColor(textBox.FillColor);
        OfficeColor? outline = ParseOfficeColor(textBox.OutlineColor);
        if (!fill.HasValue && !outline.HasValue) {
            return;
        }

        OfficeShape frame = OfficeShape.Rectangle(width, height);
        frame.FillColor = fill;
        if (textBox.FillTransparency.HasValue && textBox.FillTransparency.Value > 0 && textBox.FillTransparency.Value < 100) {
            frame.FillOpacity = 1D - textBox.FillTransparency.Value / 100D;
        }

        frame.StrokeColor = outline;
        frame.StrokeWidth = outline.HasValue ? textBox.OutlineWidthPoints ?? 0.75D : 0D;
        frame.StrokeDashStyle = MapDash(textBox.OutlineDash.ToOpenXml());
        canvas.Shape(frame, x, y, rotationAngle: textBox.Rotation ?? 0D);
    }

}
