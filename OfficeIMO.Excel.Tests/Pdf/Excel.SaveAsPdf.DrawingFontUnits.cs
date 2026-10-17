using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Excel;
using OfficeIMO.Excel.Pdf;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace OfficeIMO.Tests;

public partial class Excel {
    [Theory]
    [InlineData(100)]
    [InlineData(50)]
    public void DrawingText_PreservesPointSizeAcrossPdfAndImageCoordinates(int percent) {
        string path = Path.Combine(_directoryWithFiles, "DrawingFontUnits" + percent + ".xlsx");
        using (var document = ExcelDocument.Create(path, "Shapes")) {
            document.Sheets[0].CellValue(1, 1, "Body");
            document.Sheets[0].SetPageSetup(scale: (uint)percent);
            document.Save();
        }
        AddRenderableDrawingShape(path);
        using (var package = SpreadsheetDocument.Open(path, true)) {
            var drawing = package.WorkbookPart!.WorksheetParts.Single().DrawingsPart!.WorksheetDrawing;
            var run = drawing.Descendants<Xdr.Shape>().Single().TextBody!.Descendants<A.Run>().Single();
            run.RunProperties = new A.RunProperties { FontSize = 2400 };
            drawing.Save();
        }
        using var loaded = ExcelDocument.Load(path);
        using var pdf = UglyToad.PdfPig.PdfDocument.Open(loaded.ToPdfBytes());
        var letter = Assert.Single(pdf.GetPage(1).Letters.Where(letter => letter.Value == "V"));
        Assert.InRange(letter.FontSize, 24D * percent / 100D - 0.01D, 24D * percent / 100D + 0.01D);
        // SVG geometry is in CSS pixels: authored 24 points occupy 32 pixels.
        string svg = loaded.Sheets[0].Range("A1:D5").ToSvg();
        Assert.Contains("font-size=\"32\"", svg, StringComparison.Ordinal);
    }
}
