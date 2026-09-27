using System;
using System.IO;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Drawing.Tests;

public class DrawingSystemUiFontsTests {
    [Theory]
    [InlineData(16D)]
    [InlineData(24D)]
    public void InstalledVariableBoldPaintMatchesExplicitWeightWithoutSyntheticEmboldening(double size) {
        const string path = "/System/Library/Fonts/SFNS.ttf";
        if (!File.Exists(path)) return;
        var fonts = new OfficeFontFaceCollection {
            FontVariationResolver = _ => new System.Collections.Generic.Dictionary<string, float> { ["wght"] = 700F }
        };
        fonts.Add("Reference", File.ReadAllBytes(path), new OfficeFontFaceDescriptor(700));
        byte[] Paint(string family, OfficeFontFaceCollection? scopedFonts) {
            var image = new OfficeRasterImage(300, 80);
            new OfficeRasterCanvas(image, fonts: scopedFonts).DrawText("Animation", 8D, 8D, 280D, 60D,
                OfficeColor.Black, size, OfficeTextAlignment.Left, OfficeFontStyle.Bold, family);
            return image.GetPixels();
        }
        byte[] expected = Paint("Reference", fonts);
        Assert.Contains(expected, channel => channel != 0);
        Assert.Equal(expected, Paint("system-ui", null));
    }

    [Theory]
    [InlineData("system-ui")]
    [InlineData("-apple-system")]
    [InlineData("BlinkMacSystemFont")]
    public void MacSystemUiNamesSelectSfnsForMeasurement(string family) {
        const string path = "/System/Library/Fonts/SFNS.ttf";
        if (!File.Exists(path)) return;
        Assert.NotNull(OfficeTrueTypeFont.TryLoadFontFamily(family, out string? selectedPath));
        Assert.Equal(path, selectedPath);
    }
    [Theory]
    [InlineData("system-ui")]
    [InlineData("-apple-system")]
    [InlineData("BlinkMacSystemFont")]
    public void MacSystemUiRasterMeasurementUsesAuthoredOpticalSize(string family) {
        if (!File.Exists("/System/Library/Fonts/SFNS.ttf")) return;
        var canvas = new OfficeRasterCanvas(new OfficeRasterImage(200, 80));
        Assert.Equal(72.5625D, canvas.MeasureText("Animation", 16D, family), 9);
    }

    [Fact]
    public void EquivalentInstalledAliasesShareOneScopedProgramAndByteBudget() {
        if (!File.Exists("/System/Library/Fonts/SFNS.ttf")) return;
        var fonts = new OfficeFontFaceCollection();
        Assert.True(fonts.TryAddInstalledFamily("system-ui", OfficeFontFaceDescriptor.Regular,
            32 * 1024 * 1024, default, out int firstBytes, out string? error), error);
        Assert.True(firstBytes > 0);
        Assert.True(fonts.TryAddInstalledFamily("-apple-system", OfficeFontFaceDescriptor.Regular,
            0, default, out int aliasBytes, out error), error);
        Assert.Equal(0, aliasBytes);
        Assert.True(fonts.TryAddInstalledFamily("system-ui", new OfficeFontFaceDescriptor(700),
            0, default, out int weightedBytes, out error), error);
        Assert.Equal(0, weightedBytes);
        Assert.True(fonts.TryResolveFaceForText("Animation", "system-ui", new OfficeFontFaceDescriptor(700), 16D, out var bold));
        Assert.Equal(700, bold!.Descriptor.Weight);

        Assert.True(fonts.TryResolveFaceForText("Animation", "system-ui", OfficeFontStyle.Regular, 16D, out var first));
        Assert.True(fonts.TryResolveFaceForText("Animation", "-apple-system", OfficeFontStyle.Regular, 16D, out var alias));
        Assert.Same(first!.Program, alias!.Program);
        Assert.Equal(72.5625D, alias.Program.Measure("Animation", 16D), 9);
    }

    [Fact]
    public void InstallingAliasWeightsInDifferentOrdersKeepsEveryDescriptor() {
        if (!File.Exists("/System/Library/Fonts/SFNS.ttf")) return;
        var fonts = new OfficeFontFaceCollection();
        string[] families = { "system-ui", "-apple-system", "BlinkMacSystemFont" };
        int bytes = 0;
        foreach (var request in new[] { (families[0], 400), (families[1], 700), (families[1], 400),
            (families[2], 700), (families[0], 700), (families[2], 400) }) {
            Assert.True(fonts.TryAddInstalledFamily(request.Item1, new OfficeFontFaceDescriptor(request.Item2),
                32 * 1024 * 1024, default, out int added, out string? error), error);
            bytes += added;
        }
        Assert.Equal(File.ReadAllBytes("/System/Library/Fonts/SFNS.ttf").Length * 2, bytes);
        foreach (string family in families) {
            foreach (int weight in new[] { 400, 700 }) {
                Assert.True(fonts.TryResolveFaceForText("Animation", family, new OfficeFontFaceDescriptor(weight), 16D, out var face));
                Assert.Equal(weight, face!.Descriptor.Weight);
            }
        }
    }

    [Fact]
    public void HostVariationSelectionRunsIndependentlyForEachAlias() {
        if (!File.Exists("/System/Library/Fonts/SFNS.ttf")) return;
        var fonts = new OfficeFontFaceCollection {
            FontVariationResolver = request => new System.Collections.Generic.Dictionary<string, float> {
                ["wght"] = request.FamilyName == "system-ui" ? 500F : 600F
            }
        };
        Assert.True(fonts.TryAddInstalledFamily("system-ui", OfficeFontFaceDescriptor.Regular, 32 * 1024 * 1024,
            default, out _, out string? error), error);
        Assert.True(fonts.TryAddInstalledFamily("-apple-system", OfficeFontFaceDescriptor.Regular, 32 * 1024 * 1024,
            default, out _, out error), error);
        Assert.True(fonts.TryResolveFaceForText("Animation", "system-ui", OfficeFontStyle.Regular, 16D, out var first));
        Assert.True(fonts.TryResolveFaceForText("Animation", "-apple-system", OfficeFontStyle.Regular, 16D, out var second));
        Assert.NotEqual(first!.Program.Fingerprint, second!.Program.Fingerprint);
    }

    [Theory]
    [InlineData(OfficeFontStyle.Regular, 72.5625D)]
    [InlineData(OfficeFontStyle.Bold, 79.0390625D)]
    public void SvgAndRasterUiFallbackUseTheSameSelectedSizeAndWeight(OfficeFontStyle style, double width) {
        if (!File.Exists("/System/Library/Fonts/SFNS.ttf")) return;
        var canvas = new OfficeRasterCanvas(new OfficeRasterImage(200, 80));
        Assert.Equal(width, canvas.MeasureText("Animation", 16D, "system-ui", style), 9);
        string svg = "<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 200 80'><text x='8' y='30' "
            + "font-family='system-ui' font-size='16' font-weight='" + (style == OfficeFontStyle.Bold ? 700 : 400)
            + "'>Animation</text></svg>";
        Assert.True(OfficeSvgDrawingReader.TryRead(System.Text.Encoding.UTF8.GetBytes(svg), out OfficeDrawing? drawing, out int unsupported));
        Assert.Equal(0, unsupported);
        OfficeDrawingText text = Assert.Single(drawing!.Elements.OfType<OfficeDrawingText>());
        Assert.Equal(width, text.Width, 9);
    }
}
