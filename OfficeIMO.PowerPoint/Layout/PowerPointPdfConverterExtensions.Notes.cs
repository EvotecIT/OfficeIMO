using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.PowerPoint.Pdf;

public static partial class PowerPointPdfConverterExtensions {
    private static void RenderSpeakerNotes(PdfCore.PdfPageCanvas canvas, string notes,
        double x, double y, double width, double height, double fontSize,
        PdfCore.PdfColor color, int slideNumber, PowerPointToPdfOptions options) {
        var style = new PdfCore.PdfCanvasTextBoxStyle {
            Background = null,
            BorderColor = null,
            BorderWidth = 0D,
            PaddingX = 0D,
            PaddingY = 0D,
            FontSize = fontSize,
            TextColor = color
        };
        canvas.TextBox(notes, x, y, width, height, style, diagnosticHandler: diagnostic =>
            AddLayoutWarning(options, slideNumber, "speaker-notes-overflow", diagnostic.Message,
                diagnostic.Kind, "PowerPointSpeakerNotes", diagnostic.Message, x, y, width, height));
    }
}
