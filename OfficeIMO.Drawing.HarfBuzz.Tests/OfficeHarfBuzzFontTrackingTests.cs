using System;
using System.IO;
using System.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.Drawing.HarfBuzz;
using OfficeIMO.TestAssets;
using Xunit;

namespace OfficeIMO.Drawing.HarfBuzz.Tests;

public sealed class OfficeHarfBuzzFontTrackingTests {
    [Fact]
    public void ProviderReturnsNominalAdvancesForOwnedSizeDependentTracking() {
        byte[] table = ManagedTextShapingTestAssets.CreateTrackingTable(new[] { 10D, 20D }, new[] { 0D },
            new[] { new short[] { -100, 0 } });
        byte[] data = ManagedTextShapingTestAssets.AddTrackingTable(
            File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "Fonts", "Carlito-Regular.ttf")), table, includeStat: true);
        IOfficeFontProgram font = Assert.Single(new OfficeFontFaceCollection().Add("Tracked", data).Faces).Program;
        var request = new OfficeTextShapingRequest("AA", "Tracked", data, false, font.UnitsPerEm,
            OfficeTextDirection.LeftToRight, null, cancellationToken: default);
        OfficeTextShapingResult run = Assert.IsType<OfficeTextShapingResult>(
            OfficeHarfBuzzTextShapingProvider.Instance.ShapeText(request));
        Assert.True(font.TryGetGlyphMetrics('A', out _, out int nominal));
        Assert.Equal(new int?[] { nominal, nominal }, run.Glyphs.Select(glyph => glyph.AdvanceWidth));
        Assert.Equal(font.Measure("AA", 15D), font.MeasureShapedText("AA", run, 15D), 9);
    }
}
