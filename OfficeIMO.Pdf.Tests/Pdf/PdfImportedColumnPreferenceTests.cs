using OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfImportedColumnPreferenceTests {
    [Theory]
    [InlineData(0, false)]
    [InlineData(1, false)]
    [InlineData(2, false)]
    [InlineData(0, true)]
    public void ImportedContentMinimumCompressionPreservesPaddingAndAllClippedCellText(int automaticColumn, bool explicitMaximum) {
        const string leftText = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
        const string rightText = "ABCDEFGHIJKLM";
        var options = new PdfOptions {
            PageWidth = explicitMaximum ? 240D : 119D, PageHeight = 300D,
            MarginLeft = 20D, MarginRight = 20D, MarginTop = 20D, MarginBottom = 20D,
            DefaultFont = PdfStandardFont.Courier
        };
        var style = new PdfTableStyle {
            HeaderRowCount = 0, AutoFitColumns = true, AutoFitWidthUsesContentMinimum = true,
            PreserveWidth = true, ClipTextToCellBounds = true,
            AllowWidthBeyondFrame = explicitMaximum, PreferredWidth = explicitMaximum ? 200D : 79D,
            MaxWidth = explicitMaximum ? 79D : null,
            AutoFitColumnPreferredWidthPoints = new List<double?> {
                automaticColumn == 1 ? null : 60D, automaticColumn == 2 ? null : 140D },
            CellPaddingX = 2D, CellPaddingY = 0D, FontSize = 10D, BorderColor = null
        };
        using var pdf = PdfPigDocument.Open(PdfDocument.Create(options)
            .Table(new[] { new[] { leftText, rightText } }, style: style).ToBytes());
        var page = Assert.Single(pdf.GetPages());
        Assert.Equal(leftText + rightText, string.Concat(page.Letters.Select(letter => letter.Value)));
        Assert.InRange(page.Letters[leftText.Length].StartBaseLine.X - page.Letters[0].StartBaseLine.X,
            51.331D, 51.335D);
        Assert.All(page.Letters, letter => Assert.InRange(letter.BoundingBox.Right, 20D, 99.001D));
    }

    [Theory]
    [InlineData(100D, 31.428571D)]
    [InlineData(400D, 90D)]
    public void ImportedPreferredFrameRetainsAnExplicitMaximumBeyondTheTextFrame(double maximum, double expectedFirstWidth) {
        var options = new PdfOptions {
            PageWidth = 120D, PageHeight = 200D,
            MarginLeft = 20D, MarginRight = 20D, MarginTop = 20D, MarginBottom = 20D
        };
        var style = new PdfTableStyle {
            HeaderRowCount = 0, AutoFitColumns = true, AutoFitWidthUsesContentMinimum = true,
            AllowWidthBeyondFrame = true, PreferredWidth = 300D, MaxWidth = maximum, PreserveWidth = true,
            AutoFitColumnPreferredWidthPoints = new List<double?> { 60D, 140D },
            ColumnMinWidthPoints = new List<double?> { 20D, 40D },
            CellPaddingX = 0D, CellPaddingY = 0D, FontSize = 10D, BorderColor = null
        };
        using var pdf = PdfPigDocument.Open(PdfDocument.Create(options)
            .Table(new[] { new[] { "A", "B" } }, style: style).ToBytes());
        var page = Assert.Single(pdf.GetPages());
        var left = Assert.Single(page.Letters, letter => letter.Value == "A");
        var right = Assert.Single(page.Letters, letter => letter.Value == "B");
        Assert.Equal(expectedFirstWidth, right.StartBaseLine.X - left.StartBaseLine.X, 3);
    }

    [Theory]
    [InlineData(100D, 300D, 0D, 0, 31.428571D)]
    [InlineData(100D, 300D, 25D, 0, 25D)]
    [InlineData(300D, 300D, 0D, 0, 90D)]
    [InlineData(300D, 100D, 0D, 0, 90D)]
    [InlineData(100D, 300D, 0D, 1, 20D)]
    [InlineData(300D, 300D, 0D, 1, 160D)]
    [InlineData(300D, 300D, 25D, 1, 25D)]
    [InlineData(100D, 300D, 0D, 2, 60D)]
    [InlineData(300D, 300D, 0D, 2, 60D)]
    public void ImportedPreferencesShrinkAfterMinimumsAndGrowProportionally(
        double preferredWidth, double frameWidth, double firstMaximum, int automaticColumn, double expectedFirstWidth) {
        var options = new PdfOptions {
            PageWidth = frameWidth + 40D, PageHeight = 200D,
            MarginLeft = 20D, MarginRight = 20D, MarginTop = 20D, MarginBottom = 20D
        };
        var style = new PdfTableStyle {
            HeaderRowCount = 0, AutoFitColumns = true, AutoFitWidthUsesContentMinimum = true,
            AllowWidthBeyondFrame = true, PreferredWidth = preferredWidth, PreserveWidth = true,
            AutoFitColumnPreferredWidthPoints = new List<double?> {
                automaticColumn == 1 ? null : 60D, automaticColumn == 2 ? null : 140D },
            ColumnMinWidthPoints = new List<double?> { 20D, 40D },
            ColumnMaxWidthPoints = firstMaximum > 0D ? new List<double?> { firstMaximum, null } : null,
            CellPaddingX = 0D, CellPaddingY = 0D, FontSize = 10D, BorderColor = null
        };
        byte[] bytes = PdfDocument.Create(options).Table(new[] { new[] { "A", "B" } }, style: style).ToBytes();
        using var pdf = PdfPigDocument.Open(bytes);
        var page = Assert.Single(pdf.GetPages());
        var left = Assert.Single(page.Letters, letter => letter.Value == "A");
        var right = Assert.Single(page.Letters, letter => letter.Value == "B");
        Assert.Equal(expectedFirstWidth, right.StartBaseLine.X - left.StartBaseLine.X, 3);
        Assert.Equal(left.StartBaseLine.Y, right.StartBaseLine.Y, 3);
    }
}
