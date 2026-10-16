using System.Collections.Generic;
using System.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.TestAssets;
using Xunit;

namespace OfficeIMO.Drawing.Tests;

public sealed partial class DrawingTrueTypeVariableFontTests {
    [Theory]
    [InlineData(0F, 0D, 500)]
    [InlineData(0.000030517578125F, 0.015625D, 501)] // +1/32768 rounds up to +1/16384
    [InlineData(-0.000030517578125F, 0D, 500)] // -1/32768 rounds toward zero
    [InlineData(2F, 256D, 16884)] // clamp above maximum
    [InlineData(-2F, -256D, 500)] // clamp below minimum
    public void PublicVariableFontNormalizesHalfStepsAndEndpoints(float coordinate, double left, int advance) {
        AssertNormalizationFont(ManagedTextShapingTestAssets.CreateNormalizationFont(), coordinate, left, advance);
    }

    [Fact]
    public void PublicVariableFontNormalizesGvarPhantomMetricsWithItsContours() {
        AssertNormalizationFont(ManagedTextShapingTestAssets.CreateNormalizationFont(hvar: false),
            0.000030517578125F, 0.015625D, 501);
    }

    [Theory]
    [InlineData(0.00000762939453125F, 0.03125D, 502)] // positive 16.16 input half-step
    [InlineData(-0.00000762939453125F, 0D, 500)] // negative half-step takes the next higher value
    public void PublicVariableFontConvertsFloatInputBeforeNormalization(float coordinate, double left, int advance) {
        AssertNormalizationFont(ManagedTextShapingTestAssets.CreateNormalizationFont(extent: 0.125D),
            coordinate, left, advance);
    }

    [Theory]
    [InlineData(300F, 153.59375D, 10330)]
    [InlineData(-300F, -153.59375D, 500)]
    public void PublicVariableFontTruncatesIntermediateFixedDivision(float coordinate, double left, int advance) {
        // 300/500 truncates to 39321 in 16.16 before the prescribed 2.14 conversion.
        AssertNormalizationFont(ManagedTextShapingTestAssets.CreateNormalizationFont(extent: 500D),
            coordinate, left, advance);
    }

    [Theory]
    [InlineData(0.000030517578125F, 0.015625D, 501)] // avar maps two 16.16 units to three before 2.14 conversion
    [InlineData(-0.000030517578125F, -0.015625D, 500)]
    [InlineData(0.0000152587890625F, 0D, 500)] // truncate the intermediate avar product before 2.14 rounding
    [InlineData(0.5F, 192D, 12788)]
    public void PublicVariableFontAppliesAvarBeforeFinalCoordinateRounding(float coordinate, double left, int advance) {
        AssertNormalizationFont(ManagedTextShapingTestAssets.CreateNormalizationFont(avar: true),
            coordinate, left, advance);
    }

    [Fact]
    public void VariableFontProviderAndCacheIdentityUseTheSelectedFixedInput() {
        byte[] data = ManagedTextShapingTestAssets.CreateNormalizationFont(extent: 0.125D);
        var provider = new CapturingFontProgramProvider();
        var fonts = new OfficeFontFaceCollection {
            FontVariationResolver = _ => new Dictionary<string, float> { ["TEST"] = 0.00000762939453125F },
            FontProgramProvider = provider
        };
        Assert.True(fonts.TryAdd("Normalization", data));

        Assert.Equal(0.0000152587890625F, provider.LastRequest!.VariationCoordinates["TEST"]);
        OfficeFontFace equivalent = Load(data,
            new Dictionary<string, float> { ["TEST"] = 0.0000152587890625F }, "Normalization");
        Assert.Equal(equivalent.Program.Fingerprint, Assert.Single(fonts.Faces).Program.Fingerprint);
    }

    private static void AssertNormalizationFont(byte[] data, float coordinate, double left, int advance) {
        OfficeFontFace face = Load(data, new Dictionary<string, float> { ["TEST"] = coordinate }, "Normalization");
        IOfficeFontProgram program = face.Program;
        Assert.Equal(left, program.GetTextContours("A", 0D, 0D, program.UnitsPerEm)
            .SelectMany(contour => contour).Min(point => point.X), 10);
        Assert.True(program.TryGetGlyphMetrics('A', out _, out int selectedAdvance));
        Assert.Equal(advance, selectedAdvance);
        Assert.Equal(advance, program.Measure("A", program.UnitsPerEm), 10);
    }
}
