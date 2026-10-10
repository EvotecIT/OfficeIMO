using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using OfficeIMO.TestAssets;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfListMarkerFeatureTests {
    [Theory]
    [InlineData(false, 0, 10D)]
    [InlineData(true, 0, 10D)]
    [InlineData(false, 1, 9.2D)]
    [InlineData(true, 1, 9.2D)]
    [InlineData(false, -1, 10D)]
    [InlineData(true, -1, 10D)]
    public void MarkerFeaturesRemainConsistentInFlowAndRowColumns(bool inRow, int kerning, double advance) {
        var options = new PdfOptions { DefaultFont = PdfStandardFont.Courier, DefaultFontSize = 10D };
        options.RegisterNamedFontFamily(new PdfEmbeddedFontFamily(ManagedTextShapingTestAssets.FamilyName,
            ManagedTextShapingTestAssets.CreateFontWithPairPositioning('A', 'V')));
        var style = new PdfListStyle {
            MarkerFontFamily = ManagedTextShapingTestAssets.FamilyName,
            MarkerFontSize = 20D,
            MarkerWidth = 30D,
            MarkerGap = 12D,
            MarkerAlign = PdfAlign.Right,
            MarkerFeatureSettings = kerning < 0 ? null : OfficeTextFeatureSettings.Default.With("kern", kerning)
        };
        var items = new[] { new PdfListItem("body", marker: "AV") };
        var document = PdfDocument.Create(options);
        if (inRow) {
            document.Compose(compose => compose.Page(page => page.Content(content =>
                content.Row(row => row.PercentColumn(100, column => column.RichBullets(items, style: style))))));
        } else {
            document.RichBullets(items, style: style);
        }
        // List blocks snapshot their style when composed, including marker features.
        style.MarkerFeatureSettings = OfficeTextFeatureSettings.Default.With("kern", kerning < 0 ? 0 : 1 - kerning);
        using var pdf = PdfPigDocument.Open(document.ToBytes());
        var letters = pdf.GetPage(1).Letters;
        var a = Assert.Single(letters, letter => letter.Value == "A");
        var v = Assert.Single(letters, letter => letter.Value == "V");
        var body = Assert.Single(letters, letter => letter.Value == "b");
        Assert.Equal(advance, v.StartBaseLine.X - a.StartBaseLine.X, 3);
        // This fixture adjusts both placement and advance: with GPOS enabled,
        // its final pen is 0.2 points left of PdfPig's nominal final glyph end.
        Assert.Equal(kerning == 1 ? 11.8D : 12D, body.StartBaseLine.X - v.EndBaseLine.X, 3);
    }
}
