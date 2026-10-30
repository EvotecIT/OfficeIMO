using OfficeIMO.Drawing;
using PdfCore = OfficeIMO.Pdf;
using PptCore = OfficeIMO.PowerPoint;

namespace OfficeIMO.PowerPoint.Pdf;

public static partial class PowerPointPdfConverterExtensions {
    /// <summary>Projects legacy table previews through the same table layout as slide exports.</summary>
    internal static OfficeDrawing CreateTableDrawingForExport(PptCore.PowerPointTable table) {
        if (!table.TryGetBoundsPoints(out _, out _, out double width, out double height) || width <= 0D || height <= 0D)
            throw new InvalidOperationException("The table has no positive renderable bounds.");
        var options = new PowerPointPdfExportOptions();
        // Master/layout tables have a relationship owner but no consuming slide.
        // Discover only this table's fonts, using the same registration path as
        // slide export, before creating the canonical layout measurer.
        var pdfOptions = new PdfCore.PdfCreateOptions { Margins = PdfCore.PdfPageMargins.Uniform(OfficeIMO.OfficeLength.Points(0)) };
        pdfOptions.UseContentStreamCompressionByDefault();
        pdfOptions.ReportDiagnosticsTo(options.Report, "OfficeIMO.PowerPoint.Pdf");
        var families = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        HashSet<PdfCore.PdfStandardFont> slots = pdfOptions.CreateRegisteredFontFamilySlots(includeDocumentFontSlots: false);
        RegisterPresentationTableFonts(table, 1, pdfOptions, families, slots, options);
        RegisterPresentationFontCandidate(ResolveTableDefaultFontFamily(table), pdfOptions, families, slots, options, 1);
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
        PdfCore.PdfDocument document = PdfCore.PdfDocument.CreateForImportedLayout(pdfOptions);
        document.Canvas(target => target.AddItems(canvas.Items));
        return PdfCore.PdfReadDocument.Open(document.ToBytes()).Pages[0].ToDrawing();
    }
}
