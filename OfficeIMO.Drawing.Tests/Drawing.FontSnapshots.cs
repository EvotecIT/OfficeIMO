using OfficeIMO.Drawing;
using OfficeIMO.TestAssets;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class DrawingFontSnapshotTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void DrawingSnapshotRetainsScopedGlyphsWhenClonedOrNested(bool nested) {
        byte[] font = ManagedTextShapingTestAssets.CreateFontWithDistinctGlyphs(65, 66, 67);
        var source = new OfficeDrawing(120D, 40D)
            .AddFont("Snapshot face", font)
            .AddText("ABC", 4D, 4D, 100D, 30D, new OfficeFontInfo("Snapshot face", 20D), OfficeColor.Black);
        OfficeDrawing copy = nested
            ? new OfficeDrawing(120D, 40D).AddEffectDrawing(source, OfficeTransform.Identity)
            : source.Clone();
        Assert.Equal(OfficeDrawingRasterRenderer.ToPng(source), OfficeDrawingRasterRenderer.ToPng(copy));
        Assert.Single(copy.Fonts.Faces);
        source.Fonts.Add("Second face", font);
        Assert.Single(copy.Fonts.Faces);
    }
}
