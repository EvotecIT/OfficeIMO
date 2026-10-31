using System.Collections.Generic;
using System.Linq;
using OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfImportedSpacerWrappingTests {
    [Theory]
    [InlineData(4.8D, 40D)]
    [InlineData(5D, 40D)]
    [InlineData(3D, 43D)]
    public void ImportedCellSpacerDoesNotWidenRemainingChunkBudget(double spacerWidth, double firstTextX) {
        const string text = "a-a-a-a-a-a-a-a-a-a-a-a";
        var cell = new PdfTableCell(new[] {
            PdfTextRun.Inline(new PdfInlineBox(global::OfficeIMO.OfficeLength.Points(spacerWidth), global::OfficeIMO.OfficeLength.Points(0.01D), borderWidth: global::OfficeIMO.OfficeLength.Points(0D)) { IsTextSpacer = true }),
            PdfTextRun.Normal(text, fontSize: global::OfficeIMO.OfficeLength.Points(1D), font: PdfStandardFont.Helvetica)
        });
        var options = new PdfCreateOptions {
            PageWidth = global::OfficeIMO.OfficeLength.Points(200D), PageHeight = global::OfficeIMO.OfficeLength.Points(200D),
            MarginLeft = global::OfficeIMO.OfficeLength.Points(40D), MarginRight = global::OfficeIMO.OfficeLength.Points(40D), MarginTop = global::OfficeIMO.OfficeLength.Points(40D), MarginBottom = global::OfficeIMO.OfficeLength.Points(40D)
        };
        var style = new PdfTableStyle {
            HeaderRowCount = 0, ColumnWidths = new List<global::OfficeIMO.OfficeLength?> { global::OfficeIMO.OfficeLength.Points(5D) },
            CellPaddingX = global::OfficeIMO.OfficeLength.Points(0D), CellPaddingY = global::OfficeIMO.OfficeLength.Points(0D), FontSize = global::OfficeIMO.OfficeLength.Points(1D)
        };
        byte[] bytes = PdfDocument.Create(options).Table(new[] { new[] { cell } }, style: style).ToBytes();
        using var pdf = PdfPigDocument.Open(bytes);
        var letters = pdf.GetPage(1).Letters.Where(letter => letter.Value == "a" || letter.Value == "-").ToArray();
        Assert.Equal(text, string.Concat(letters.Select(letter => letter.Value)));
        Assert.Equal(firstTextX, letters[0].StartBaseLine.X, 3);
        Assert.All(letters, letter => Assert.InRange(letter.StartBaseLine.X + letter.Width, 40D, 45.001D));
    }
}
