using OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfImportedMixedPercentageTests {
    [Theory]
    [InlineData(1D, false)]
    [InlineData(1.5D, false)]
    [InlineData(1D, true)]
    [InlineData(1.5D, true)]
    public void AutomaticFramesWithOvercommittedMixedPercentagesRemainFinite(double fraction, bool autoFit) {
        var style = MixedStyle(6.4D);
        style.PreferredWidth = null;
        style.AllowWidthBeyondFrame = true;
        style.AutoFitUnspecifiedWidthToContent = true;
        style.AutoFitColumns = autoFit;
        style.ImportedColumnWidthFractions = new List<double?> { fraction, null, null };
        var (first, second) = RenderWidths(style);
        Assert.InRange(first, 280D, 288D);
        Assert.InRange(second, 6.39D, 10D);
    }

    [Fact]
    public void OvercommittedAutomaticFramesReserveGapsWithinTheirAvailableExtent() {
        var style = MixedStyle(6.4D);
        style.PreferredWidth = null;
        style.AllowWidthBeyondFrame = true;
        style.AutoFitUnspecifiedWidthToContent = true;
        style.CellSpacing = global::OfficeIMO.OfficeLength.Points(10D);
        style.ImportedColumnWidthFractions = new List<double?> { 1D, null, null };
        var (first, second) = RenderWidths(style);
        Assert.InRange(first + second, 290D, 294D);
    }

    [Theory]
    [InlineData(.25D, 50.35D, 75D, 75D)]
    [InlineData(.5D, 50.35D, 150D, 59.6D)]
    [InlineData(.75D, 50.35D, 199.3D, 50.35D)]
    [InlineData(.25D, 6.4D, 75D, 75D)]
    [InlineData(.5D, 6.4D, 150D, 50.7D)]
    [InlineData(.75D, 6.4D, 225D, 26.5D)]
    public void MixedPercentageColumnsReserveMinimumsAndCompressRemainingPreferences(
        double fraction, double minimum, double firstWidth, double secondWidth) {
        var style = MixedStyle(minimum);
        style.ImportedColumnWidthFractions = new List<double?> { fraction, null, null };
        style.AutoFitColumnPreferredWidthPoints = new List<double?> { null, 72D, 144D };
        var (first, second) = RenderWidths(style);
        Assert.Equal(firstWidth, first, 3);
        Assert.InRange(second, secondWidth - .05D, secondWidth + .05D);
    }

    [Theory]
    [InlineData(50.35D, 187.2375D, 62.4125D)]
    [InlineData(6.4D, 220.2D, 73.4D)]
    public void OvercommittedPercentagesReserveUnspecifiedColumnsBeforeScaling(
        double minimum, double firstWidth, double secondWidth) {
        var style = MixedStyle(minimum);
        style.ImportedColumnWidthFractions = new List<double?> { .75D, .75D, null };
        style.AutoFitColumnPreferredWidthPoints = new List<double?> { null, null, 100D };
        var (first, second) = RenderWidths(style);
        Assert.Equal(firstWidth, first, 3);
        Assert.Equal(secondWidth, second, 3);
    }

    [Fact]
    public void MixedPercentageAllocationKeepsAbsolutePreferencesBeforeGivingSpaceToAutomaticColumns() {
        var style = MixedStyle(6.4D);
        style.AutoFitColumns = true;
        style.AutoFitWidthUsesContentMinimum = true;
        style.ImportedColumnWidthFractions = new List<double?> { .25D, null, null };
        style.AutoFitColumnPreferredWidthPoints = new List<double?> { null, 72D, null };
        var (first, second) = RenderWidths(style);
        Assert.Equal(75D, first, 3);
        Assert.Equal(72D, second, 3);
    }

    private static PdfTableStyle MixedStyle(double minimum) => new() {
        HeaderRowCount = 0, PreferredWidth = global::OfficeIMO.OfficeLength.Points(300D), PreserveWidth = true,
        ColumnMinWidths = new List<global::OfficeIMO.OfficeLength?> { global::OfficeIMO.OfficeLength.Points(minimum), global::OfficeIMO.OfficeLength.Points(minimum), global::OfficeIMO.OfficeLength.Points(minimum) },
        CellPaddingX = global::OfficeIMO.OfficeLength.Points(0D), CellPaddingY = global::OfficeIMO.OfficeLength.Points(0D), FontSize = global::OfficeIMO.OfficeLength.Points(10D), BorderColor = null
    };

    [Fact]
    public void FixedMixedPercentagesUseTheWholeGridFallbackWhenPhysicalMinimumsExceedTheFrame() {
        var style = MixedStyle(11.8D);
        style.PreferredWidth = global::OfficeIMO.OfficeLength.Points(20D);
        style.FontSize = global::OfficeIMO.OfficeLength.Points(5D);
        style.ImportedColumnWidthFractions = new List<double?> { .5D, null, null };
        style.AutoFitColumnPreferredWidthPoints = new List<double?> { null, 72D, 100D };
        var (first, second) = RenderWidths(style);
        Assert.InRange(first, 20D / 3D - .002D, 20D / 3D + .002D);
        Assert.InRange(second, 20D / 3D - .002D, 20D / 3D + .002D);
    }

    private static (double First, double Second) RenderWidths(PdfTableStyle style) {
        var options = new PdfCreateOptions {
            PageWidth = global::OfficeIMO.OfficeLength.Points(340D), PageHeight = global::OfficeIMO.OfficeLength.Points(200D),
            MarginLeft = global::OfficeIMO.OfficeLength.Points(20D), MarginRight = global::OfficeIMO.OfficeLength.Points(20D), MarginTop = global::OfficeIMO.OfficeLength.Points(20D), MarginBottom = global::OfficeIMO.OfficeLength.Points(20D)
        };
        using var pdf = PdfPigDocument.Open(PdfDocument.Create(options)
            .Table(new[] { new[] { "A", "B", "C" } }, style: style).ToBytes());
        var page = Assert.Single(pdf.GetPages());
        var letters = page.Letters;
        Assert.Equal("ABC", string.Concat(letters.Select(letter => letter.Value)));
        double first = Assert.Single(letters, letter => letter.Value == "A").StartBaseLine.X;
        double second = Assert.Single(letters, letter => letter.Value == "B").StartBaseLine.X;
        double third = Assert.Single(letters, letter => letter.Value == "C").StartBaseLine.X;
        return (second - first, third - second);
    }
}
