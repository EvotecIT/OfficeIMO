using OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfTableTextLayoutReuseSpacingTests {
    [Theory]
    [InlineData(50D, 0D, false)]
    [InlineData(50D, 0D, true)]
    [InlineData(100D, 1D, false)]
    [InlineData(100D, 1D, true)]
    [InlineData(50D, 1D, false)]
    [InlineData(50D, 1D, true)]
    public void RepeatedCellTextPreservesEachRunsScalingAndSpacing(double scale, double spacing, bool styledFirst) {
        var natural = PdfTextRun.Normal("ABCD");
        var styled = PdfTextRun.Normal("ABCD").WithHorizontalTextScaling(scale).WithCharacterSpacing(spacing);
        byte[] bytes = PdfDocument.Create(pdf => pdf.Content(content => content.Table(new[] {
            new[] { PdfTableCell.RichTextCell(new[] { styledFirst ? styled : natural }) },
            new[] { PdfTableCell.RichTextCell(new[] { styledFirst ? natural : styled }) }
        }, style: new PdfTableStyle {
            FontSize = 12D,
            HeaderRowCount = 0,
            ColumnWidthPoints = new() { 200D },
            CellPaddingX = 0D,
            CellPaddingY = 0D
        })), new PdfCreateOptions { DefaultFont = PdfStandardFont.Courier, DefaultFontSize = global::OfficeIMO.OfficeLength.Points(12D) }).ToBytes();

        using var document = PdfPigDocument.Open(bytes);
        var rows = document.GetPage(1).Letters
            .Where(letter => !string.IsNullOrWhiteSpace(letter.Value))
            .GroupBy(letter => Math.Round(letter.StartBaseLine.Y, 2))
            .OrderByDescending(row => row.Key)
            .ToArray();
        Assert.Equal(2, rows.Length);
        for (int index = 0; index < rows.Length; index++) {
            var letters = rows[index].OrderBy(letter => letter.StartBaseLine.X).ToArray();
            Assert.Equal("ABCD", string.Concat(letters.Select(letter => letter.Value)));
            bool isStyled = index == (styledFirst ? 0 : 1);
            double expectedAdvance = 3D * (7.2D * (isStyled ? scale / 100D : 1D) + (isStyled ? spacing : 0D));
            Assert.Equal(expectedAdvance, letters[3].StartBaseLine.X - letters[0].StartBaseLine.X, 3);
        }
    }
}
