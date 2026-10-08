using System;
using System.Collections.Generic;
using System.IO;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Drawing.Tests;

public sealed class DrawingFontOpticalSizeTests {
    private const string Text = "Variable OfficeIMO";

    [Theory]
    [InlineData(10D)]
    [InlineData(16D)]
    [InlineData(24D)]
    [InlineData(48D)]
    [InlineData(200D)]
    public void AuthoredSizeMatchesAnExplicitOpticalInstanceAndPreservesWeight(double size) {
        OfficeFontFaceCollection automatic = Fonts(null);
        OfficeFontFaceCollection reference = Fonts((float)size);
        Assert.True(automatic.TryResolveFaceForText(Text, "Probe", OfficeFontStyle.Regular, size, out OfficeFontFace? selected));
        OfficeFontFace expected = Assert.Single(reference.Faces);
        Assert.Equal(expected.Program.Fingerprint, selected!.Program.Fingerprint);
        Assert.False(selected.CanEmbedAsStaticPdfFont);
        Assert.True(automatic.TryMeasureText(Text, size, "Probe", OfficeFontStyle.Regular, out double measured));
        Assert.Equal(expected.Program.Measure(Text, size), measured);
        Assert.Equal(expected.Program.GetTextContours(Text, 0, 0, size), selected.Program.GetTextContours(Text, 0, 0, size));
        Assert.Same(selected, Resolve(automatic, size));
    }

    [Fact]
    public void ExplicitOpticalSelectionWinsOverAuthoredSize() {
        OfficeFontFaceCollection fonts = Fonts(20F);
        OfficeFontFace original = Assert.Single(fonts.Faces);
        Assert.Same(original, Resolve(fonts, 10D));
        Assert.Same(original, Resolve(fonts, 48D));
        Assert.True(fonts.TryMeasureText(Text, 48D, "Probe", OfficeFontStyle.Regular, out double width));
        Assert.Equal(original.Program.Measure(Text, 48D), width);
    }

    [Fact]
    public void ProviderOwnedTrueTypeProgramKeepsItsSelectedInstance() {
        IOfficeFontProgram program = Assert.Single(Fonts(null).Faces).Program;
        var fonts = new OfficeFontFaceCollection { FontProgramProvider = new FixedProvider(program) };
        Assert.True(fonts.TryAdd("Probe", File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "TestAssets", "RobotoFlex.ttf"))));
        foreach (OfficeFontFaceCollection snapshot in new[] { fonts, fonts.Clone() }) {
            Assert.Same(program, Resolve(snapshot, 48D).Program);
            Assert.True(snapshot.TryMeasureText(Text, 48D, "Probe", OfficeFontStyle.Regular, out double width));
            Assert.Equal(program.Measure(Text, 48D), width);
        }
    }

    [Theory]
    [InlineData(16D, 1D)]
    [InlineData(16D, 2D)]
    [InlineData(24D, 1D)]
    [InlineData(24D, 2D)]
    public void RasterOutputUsesAuthoredOpticalSizeBeforeOutputScale(double size, double scale) {
        OfficeFontFaceCollection automatic = Fonts(null);
        OfficeFontFaceCollection reference = Fonts((float)size);
        byte[] Render(OfficeFontFaceCollection fonts) {
            Assert.True(fonts.TryMeasureText(Text, size, "Probe", OfficeFontStyle.Regular, out double width));
            var drawing = new OfficeDrawing(500D, 100D).AddPositionedText(Text, 10D, 10D, width, 60D,
                new OfficeFontInfo("Probe", size), OfficeColor.Black, textAdvanceWidth: width);
            drawing.Fonts.AddRange(fonts);
            return OfficeDrawingRasterRenderer.Render(drawing, scale, OfficeColor.White).GetPixels();
        }
        Assert.Equal(Render(reference), Render(automatic));
    }

    [Theory]
    [InlineData(0D)]
    [InlineData(-1D)]
    [InlineData(double.NaN)]
    [InlineData(double.PositiveInfinity)]
    public void SizeAwareFaceResolutionRejectsInvalidSizes(double size) {
        Assert.Throws<ArgumentOutOfRangeException>(() => Fonts(null).TryResolveFaceForText(
            Text, "Probe", OfficeFontStyle.Regular, size, out _));
    }

    private static OfficeFontFace Resolve(OfficeFontFaceCollection fonts, double size) {
        Assert.True(fonts.TryResolveFaceForText(Text, "Probe", OfficeFontStyle.Regular, size, out OfficeFontFace? face));
        return face!;
    }

    private static OfficeFontFaceCollection Fonts(float? opticalSize) {
        var axes = new Dictionary<string, float> { ["wght"] = 700F };
        if (opticalSize.HasValue) axes.Add("opsz", opticalSize.Value);
        var fonts = new OfficeFontFaceCollection { FontVariationResolver = _ => axes };
        Assert.True(fonts.TryAdd("Probe", File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "TestAssets", "RobotoFlex.ttf"))));
        return fonts;
    }

    private sealed class FixedProvider(IOfficeFontProgram program) : IOfficeFontProgramProvider {
        public OfficeFontProgramLoadResult TryLoad(OfficeFontProgramLoadRequest request) =>
            new(program, request.Data.Length);
    }
}
