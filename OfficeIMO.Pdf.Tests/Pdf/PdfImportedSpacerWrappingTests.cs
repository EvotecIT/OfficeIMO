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
            PdfTextRun.Inline(new PdfInlineBox(spacerWidth, 0.01D, borderWidth: 0D) { IsTextSpacer = true }),
            PdfTextRun.Normal(text, fontSize: 1D, font: PdfStandardFont.Helvetica)
        });
        var options = new PdfOptions {
            PageWidth = 200D, PageHeight = 200D,
            MarginLeft = 40D, MarginRight = 40D, MarginTop = 40D, MarginBottom = 40D
        };
        var style = new PdfTableStyle {
            HeaderRowCount = 0, ColumnWidthPoints = new List<double?> { 5D },
            CellPaddingX = 0D, CellPaddingY = 0D, FontSize = 1D
        };
        byte[] bytes = PdfDocument.Create(options).Table(new[] { new[] { cell } }, style: style).ToBytes();
        using var pdf = PdfPigDocument.Open(bytes);
        var letters = pdf.GetPage(1).Letters.Where(letter => letter.Value == "a" || letter.Value == "-").ToArray();
        Assert.Equal(text, string.Concat(letters.Select(letter => letter.Value)));
        Assert.Equal(firstTextX, letters[0].StartBaseLine.X, 3);
        Assert.All(letters, letter => Assert.InRange(letter.StartBaseLine.X + letter.Width, 40D, 45.001D));
    }
}
