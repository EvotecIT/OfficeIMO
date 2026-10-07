using OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfRunSpacingTableSizingTests {
    [Theory]
    [InlineData(200D, 1D, 9.583333D, true)]
    [InlineData(50D, 1D, 30D, true)]
    [InlineData(200D, -1D, 11.25D, true)]
    [InlineData(200D, 1D, 9.583333D, false)]
    [InlineData(50D, 1D, 30D, false)]
    [InlineData(200D, -1D, 11.25D, false)]
    public void TableRunShrinkingUsesScaledGlyphsAndFixedTracking(double scale, double tracking, double expectedSize, bool explicitSize) {
        var options = new PdfOptions { DefaultFont = PdfStandardFont.Courier, MarginLeft = 36D, MarginRight = 36D };
        PdfTextRun run = new PdfTextRun("MMMMMMMM", font: PdfStandardFont.Courier, fontSize: explicitSize ? 30D : null)
            .WithHorizontalTextScaling(scale).WithCharacterSpacing(tracking);
        byte[] bytes = PdfDocument.Create(options).Table(new[] {
            new[] { PdfTableCell.RichTextCell(new[] { run }) }
        }, style: new PdfTableStyle {
            FontSize = 30D, MinimumShrinkFontSize = 1D, ShrinkTextToFit = true,
            ColumnWidthPoints = new() { 100D }, CellPaddingX = 0D, CellPaddingY = 0D, HeaderRowCount = 0
        }).ToBytes();
        using var pdf = PdfPigDocument.Open(bytes);
        var letters = pdf.GetPage(1).Letters;
        Assert.Equal("MMMMMMMM", string.Concat(letters.Select(letter => letter.Value)));
        Assert.All(letters, letter => {
            Assert.InRange(Math.Abs(letter.FontSize - expectedSize), 0D, 0.02D);
            Assert.Equal(letters[0].StartBaseLine.Y, letter.StartBaseLine.Y, 3);
        });
        double advance = letters[7].StartBaseLine.X - letters[0].StartBaseLine.X + letters[7].FontSize * 0.6D * scale / 100D + tracking;
        Assert.InRange(advance, 0D, 100.02D);
        // Negative tracking leaves the last glyph's ink beyond its trailing advance.
        Assert.InRange(letters.Max(letter => letter.BoundingBox.Right) - letters.Min(letter => letter.BoundingBox.Left),
            0D, 100.02D + Math.Max(0D, -tracking));
    }
}
