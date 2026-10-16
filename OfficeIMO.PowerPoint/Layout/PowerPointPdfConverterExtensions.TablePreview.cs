using OfficeIMO.Drawing;
using PdfCore = OfficeIMO.Pdf;
using PptCore = OfficeIMO.PowerPoint;

namespace OfficeIMO.PowerPoint.Pdf;

public static partial class PowerPointPdfConverterExtensions {
    /// <summary>Projects legacy table previews through the same table layout as slide exports.</summary>
    internal static OfficeDrawing CreateTableDrawingForExport(PptCore.PowerPointTable table) {
        var presentation = table.OwnerSlide?.OwnerPresentation
            ?? throw new InvalidOperationException("The table is not attached to a presentation.");
        if (!table.TryGetBoundsPoints(out _, out _, out double width, out double height) || width <= 0D || height <= 0D)
            throw new InvalidOperationException("The table has no positive renderable bounds.");
        var options = new PowerPointToPdfOptions();
        PdfCore.PdfOptions pdfOptions = CreatePdfOptions(presentation, options);
        options.NativeTextMeasure = PdfCore.PdfWriter.CreateDrawingTextMeasure(pdfOptions);
        var canvas = new PdfCore.PdfPageCanvas();
        RenderTable(canvas, table, 0D, 0D, width, height, 1, options);
        // The canonical table pass may grow rows beyond the source frame.
        // Size the preview page from that fitted scene before the PDF writer
        // validates its bounds; recreating the scene would fit it a second time.
        foreach (PdfCore.PdfCanvasTableItem item in canvas.Items.OfType<PdfCore.PdfCanvasTableItem>()) {
            width = Math.Max(width, item.X + item.Width);
            height = Math.Max(height, item.Y + item.Height);
        }
        pdfOptions.PageWidth = width;
        pdfOptions.PageHeight = height;
        PdfCore.PdfDocument document = PdfCore.PdfDocument.Create(pdfOptions);
        document.Canvas(target => target.AddItems(canvas.Items));
        return PdfCore.PdfReadDocument.Open(document.ToBytes()).Pages[0].ToDrawing();
    }
}
