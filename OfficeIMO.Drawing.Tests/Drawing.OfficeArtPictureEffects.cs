using OfficeIMO.Drawing;
using OfficeIMO.Drawing.Binary;
using Xunit;

namespace OfficeIMO.Tests;

public partial class DrawingTests {
    [Theory]
    [InlineData(-32768, 0, 0, 0)]
    [InlineData(32768, 255, 255, 255)]
    [InlineData(-16384, 10, 50, 110)]
    [InlineData(16384, 138, 178, 238)]
    public void OfficeArtPictureEffects_ProjectBrightnessEndpointsAndRetainAlpha(int brightness, byte red, byte green, byte blue) {
        var source = new OfficeRasterImage(1, 1, OfficeColor.FromRgba(20, 100, 220, 96));
        OfficeRasterImage output = PictureEffects((0x0109, unchecked((uint)brightness))).Apply(source, default);
        Assert.Equal(OfficeColor.FromRgba(red, green, blue, 96), output.GetPixel(0, 0));
        Assert.Equal(OfficeColor.FromRgba(20, 100, 220, 96), source.GetPixel(0, 0));
    }

    [Theory]
    [InlineData(0U, 128, 128, 128)]
    [InlineData(65536U, 20, 100, 220)]
    [InlineData(131072U, 0, 72, 255)]
    [InlineData(2147483647U, 0, 0, 255)]
    public void OfficeArtPictureEffects_ProjectContrastRange(uint contrast, byte red, byte green, byte blue) {
        var source = new OfficeRasterImage(1, 1, OfficeColor.FromRgba(20, 100, 220, 96));
        Assert.Equal(OfficeColor.FromRgba(red, green, blue, 96),
            PictureEffects((0x0108, contrast)).Apply(source, default).GetPixel(0, 0));
    }

    [Theory]
    [InlineData(0x00040004U, 54, 182)]
    [InlineData(0x00020002U, 0, 255)]
    [InlineData(0x00060006U, 0, 255)]
    public void OfficeArtPictureEffects_ProjectDisplayModeAndBiLevelPrecedence(uint flags, byte redGray, byte greenGray) {
        var source = new OfficeRasterImage(2, 1);
        source.SetPixel(0, 0, OfficeColor.FromRgba(255, 0, 0, 48));
        source.SetPixel(1, 0, OfficeColor.FromRgba(0, 255, 0, 96));
        OfficeRasterImage output = PictureEffects((0x013F, flags)).Apply(source, default);
        Assert.Equal(OfficeColor.FromRgba(redGray, redGray, redGray, 48), output.GetPixel(0, 0));
        Assert.Equal(OfficeColor.FromRgba(greenGray, greenGray, greenGray, 96), output.GetPixel(1, 0));
    }

    [Fact]
    public void OfficeArtPictureEffects_KeyMatchesSourceRgbBeforeToneAndLeavesNonmatchingAlpha() {
        var source = new OfficeRasterImage(2, 1, OfficeColor.FromRgba(20, 100, 220, 96));
        source.SetPixel(1, 0, OfficeColor.FromRgba(21, 100, 220, 48));
        OfficeRasterImage output = PictureEffects((0x0107, 0x00DC6414), (0x0109, 32768)).Apply(source, default);
        Assert.Equal(OfficeColor.FromRgba(255, 255, 255, 0), output.GetPixel(0, 0));
        Assert.Equal(OfficeColor.FromRgba(255, 255, 255, 48), output.GetPixel(1, 0));
        Assert.Equal(OfficeColor.FromRgba(20, 100, 220, 96), source.GetPixel(0, 0));
    }

    [Theory]
    [InlineData(0x00400040U, true)]
    [InlineData(0x00400000U, false)]
    [InlineData(0x00000040U, null)]
    public void OfficeArtPictureProperties_PreserveGraysRequiresItsUseBit(uint flags, bool? value) {
        OfficeArtPictureProperties picture = OfficeArtPictureProperties.Decode(new[] { new OfficeArtProperty(0, 0x013F, flags) });
        Assert.Equal(value, picture.PreserveGrays);
    }

    [Fact]
    public void OfficeArtPictureEffects_PreserveSourceGraysDuringToneChanges() {
        var source = new OfficeRasterImage(2, 1, OfficeColor.FromRgb(90, 90, 90));
        source.SetPixel(1, 0, OfficeColor.FromRgb(90, 20, 30));
        OfficeRasterImage output = PictureEffects((0x013F, 0x00400040), (0x0109, 32768)).Apply(source, default);
        Assert.Equal(OfficeColor.FromRgb(90, 90, 90), output.GetPixel(0, 0));
        Assert.Equal(OfficeColor.White, output.GetPixel(1, 0));
    }

    [Theory]
    [InlineData(0x0108, 0xFFFFFFFFU, (int)OfficeArtPictureEffectLimit.InvalidContrast)]
    [InlineData(0x0109, 0xFFFF7FFFU, (int)OfficeArtPictureEffectLimit.InvalidBrightness)]
    [InlineData(0x0109, 32769U, (int)OfficeArtPictureEffectLimit.InvalidBrightness)]
    [InlineData(0x011A, 0x00030201U, (int)OfficeArtPictureEffectLimit.UnqualifiedRecolor)]
    [InlineData(0x011B, 0U, (int)OfficeArtPictureEffectLimit.ExtendedColor)]
    [InlineData(0x0117, 0x20000001U, (int)OfficeArtPictureEffectLimit.ExtendedColor)]
    [InlineData(0x011D, 0x20000001U, (int)OfficeArtPictureEffectLimit.ExtendedColor)]
    public void OfficeArtPictureEffects_RetainPreciseLimitsInsteadOfGuessing(ushort property, uint value, int limit) {
        OfficeArtPictureEffectProjector effects = PictureEffects((property, value));
        Assert.False(effects.HasProjection);
        Assert.Equal((OfficeArtPictureEffectLimit)limit, effects.Limits);
    }

    [Fact]
    public void OfficeArtPictureEffects_DoesNotGuessAnUnresolvedPaletteKey() {
        OfficeArtPictureEffectProjector effects = PictureEffects((0x0107, 0x08000003));
        Assert.False(effects.HasProjection);
        Assert.Equal(OfficeArtPictureEffectLimit.UnresolvedTransparentColor, effects.Limits);
    }

    [Theory]
    [InlineData(0x0107, 0xFFFFFFFFU)]
    [InlineData(0x011A, 0xFFFFFFFFU)]
    [InlineData(0x0108, 65536U)]
    [InlineData(0x0109, 0U)]
    [InlineData(0x013F, 6U)]
    [InlineData(0x0115, 0xFFFFFFFFU)]
    [InlineData(0x011B, 0xFFFFFFFFU)]
    [InlineData(0x0117, 0x20000000U)]
    [InlineData(0x011D, 0x20000000U)]
    [InlineData(0x0116, 0xFFFFFFFFU)]
    [InlineData(0x011C, 0xFFFFFFFFU)]
    [InlineData(0x0116, 0U)]
    [InlineData(0x011C, 0U)]
    public void OfficeArtPictureEffects_DefaultOrUnusedControlsDoNotRequireRasterization(ushort property, uint value) {
        OfficeArtPictureEffectProjector effects = PictureEffects((property, value));
        Assert.False(effects.HasProjection);
        Assert.Equal(OfficeArtPictureEffectLimit.None, effects.Limits);
    }

    private static OfficeArtPictureEffectProjector PictureEffects(params (ushort Id, uint Value)[] properties) =>
        OfficeArtPictureEffectProjector.Create(OfficeArtPictureProperties.Decode(properties.Select((entry, index) =>
            new OfficeArtProperty(index, entry.Id, entry.Value)).ToArray()), reference =>
                reference.TryResolve(_ => null, out OfficeColor color) ? color : null);
}
