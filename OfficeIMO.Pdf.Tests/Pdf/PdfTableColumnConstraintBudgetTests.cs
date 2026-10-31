using System.IO;
using System.Linq;
using OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfDocumentVisualQualityTests {
    [Fact]
    public void Table_AllocatesManySimultaneouslyBindingMinimumColumns() {
        const int narrowColumns = 256;
        var options = new PdfOptions {
            PageWidth = 3060, PageHeight = 240,
            MarginLeft = 30, MarginRight = 30, MarginTop = 30, MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica, DefaultFontSize = 9
        };
        var style = TableStyles.Minimal();
        style.AutoFitColumns = false;
        style.CellPaddingX = 0;
        style.ColumnWidthWeights = Enumerable.Repeat(1D, narrowColumns).Concat(new[] { 1000D }).ToList();
        style.ColumnMinWidthPoints = Enumerable.Repeat<double?>(3D, narrowColumns).ToList();
        style.ColumnMinWidthPoints.Add(null);
        var row = Enumerable.Repeat(string.Empty, narrowColumns).Concat(new[] { "A" }).ToArray();

        byte[] bytes = PdfDocument.Create(options).Table(new[] { row }, style: style).ToBytes();

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        Assert.Contains("A", pdf.GetPage(1).Text);
    }
}
