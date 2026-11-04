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
        var options = new PdfCreateOptions {
            PageSetup = {
                Size = new PdfPageSize(OfficeLength.Points(3060), OfficeLength.Points(240)),
                Margins = { Left = OfficeLength.Points(30), Right = OfficeLength.Points(30), Top = OfficeLength.Points(30), Bottom = OfficeLength.Points(30) }
            },
            DefaultFont = PdfStandardFont.Helvetica, DefaultFontSize = OfficeLength.Points(9)
        };
        var style = PdfTableStyles.Minimal();
        style.AutoFitColumns = false;
        style.CellPaddingX = OfficeLength.Points(0);
        style.ColumnWidthWeights = Enumerable.Repeat(1D, narrowColumns).Concat(new[] { 1000D }).ToList();
        style.ColumnMinWidths = Enumerable.Repeat<OfficeLength?>(OfficeLength.Points(3D), narrowColumns).ToList();
        style.ColumnMinWidths.Add(null);
        var row = Enumerable.Repeat(string.Empty, narrowColumns).Concat(new[] { "A" }).ToArray();

        byte[] bytes = PdfDocument.Create(options).Table(new[] { row }, style: style).ToBytes();

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        Assert.Contains("A", pdf.GetPage(1).Text);
    }
}
