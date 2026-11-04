using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.PowerPoint.Pdf;

public static partial class PowerPointPdfConverterExtensions {
    private static void RenderSpeakerNotes(PdfCore.PdfPageCanvas canvas, string notes,
        double x, double y, double width, double height, double fontSize,
        global::OfficeIMO.OfficeColor color, int slideNumber, PowerPointPdfExportOptions options) {
        var style = new PdfCore.PdfCanvasTextBoxStyle {
            Background = null,
            BorderColor = null,
            BorderWidth = global::OfficeIMO.OfficeLength.Points(0D),
            PaddingX = global::OfficeIMO.OfficeLength.Points(0D),
            PaddingY = global::OfficeIMO.OfficeLength.Points(0D),
            FontSize = global::OfficeIMO.OfficeLength.Points(fontSize),
            TextColor = color
        };
        canvas.TextBox(notes, global::OfficeIMO.OfficeLength.Points(x), global::OfficeIMO.OfficeLength.Points(y), global::OfficeIMO.OfficeLength.Points(width), global::OfficeIMO.OfficeLength.Points(height), style, diagnosticHandler: diagnostic =>
            AddLayoutWarning(options, slideNumber, "speaker-notes-overflow", diagnostic.Message,
                diagnostic.Kind, "PowerPointSpeakerNotes", diagnostic.Message, x, y, width, height));
    }
}
