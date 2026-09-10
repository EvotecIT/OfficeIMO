using System.Linq;
using System.Xml.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.TestAssets;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class DrawingNestedShapingTests {
    [Theory]
    [InlineData("group")]
    [InlineData("effect")]
    [InlineData("tile")]
    [InlineData("mask")]
    [InlineData("composed")]
    public void NestedSvgUsesTheChildShapingProfile(string mode) {
        var provider = new WideGlyphProvider();
        var child = new OfficeDrawing(30, 40).AddText("AA", 0, 0, 30, 40,
            new OfficeFontInfo("Nested Font", 20), wrapText: true, shrinkToFit: true);
        child.Fonts.Add("Nested Font", ManagedTextShapingTestAssets.CreateFont('A'));
        child.ApplyImageExportOptions(new OfficeImageExportOptions { TextShapingProvider = provider, TextShapingLanguage = "ar-SA" });
        string direct = OfficeDrawingSvgExporter.ToSvg(child);
        var parent = new OfficeDrawing(30, 40).ApplyImageExportOptions(new OfficeImageExportOptions {
            TextShapingProvider = new WideGlyphProvider(400), TextShapingLanguage = "en-US"
        });
        if (mode == "effect") parent.AddEffectDrawing(child, OfficeTransform.Identity);
        else if (mode == "tile") parent.AddTilingPattern(child, new OfficeImagePlacement(0, 0, 30, 40), 30, 40);
        else if (mode == "mask") parent.AddEffectDrawing(new OfficeDrawing(30, 40), OfficeTransform.Identity,
            OfficeBlendMode.Normal, new OfficeDrawingSoftMask(child));
        else if (mode == "composed") parent.AddDrawing(new OfficeDrawing(30, 40)
            .AddClippedDrawing(child, 0, 0, OfficeClipPath.Rectangle(30, 40)), 0, 0);
        else parent.AddClippedDrawing(child, 0, 0, OfficeClipPath.Rectangle(30, 40));
        provider.Languages.Clear();
        string nested = OfficeDrawingSvgExporter.ToSvg(parent);
        Assert.Equal(TextElements(direct), TextElements(nested));
        Assert.NotEmpty(provider.Languages);
        Assert.All(provider.Languages, language => Assert.Equal("ar-SA", language));
    }

    [Theory]
    [InlineData("group")]
    [InlineData("effect")]
    [InlineData("tile")]
    public void NestedRasterKeepsChildShapingAndParentClip(string mode) {
        var child = new OfficeDrawing(30, 40).AddText("AA", 0, 0, 30, 40,
            new OfficeFontInfo("Nested Font", 20), wrapText: true, shrinkToFit: true);
        child.Fonts.Add("Nested Font", ManagedTextShapingTestAssets.CreateFont('A'));
        child.ApplyImageExportOptions(new OfficeImageExportOptions { TextShapingProvider = new WideGlyphProvider(), TextShapingLanguage = "ar-SA" });
        var inner = new OfficeDrawing(30, 40);
        if (mode == "effect") inner.AddEffectDrawing(child, OfficeTransform.Identity);
        else if (mode == "tile") inner.AddTilingPattern(child, new OfficeImagePlacement(0, 0, 30, 40), 30, 40);
        else inner.AddClippedDrawing(child, 0, 0, OfficeClipPath.Rectangle(30, 40));
        var parent = new OfficeDrawing(30, 40).ApplyImageExportOptions(new OfficeImageExportOptions {
            TextShapingProvider = new WideGlyphProvider(400), TextShapingLanguage = "en-US"
        }).AddClippedDrawing(inner, 0, 0, OfficeClipPath.Rectangle(30, 25));
        var expected = OfficeDrawingRasterRenderer.Render(child, 1, OfficeColor.White);
        var actual = OfficeDrawingRasterRenderer.Render(parent, 1, OfficeColor.White);
        for (int y = 0; y < 40; y++)
            for (int x = 0; x < 30; x++)
                Assert.Equal(y < 25 ? expected.GetPixel(x, y) : OfficeColor.White, actual.GetPixel(x, y));
    }

    [Fact]
    public void ChildProfileDoesNotLeakIntoFollowingSibling() {
        var parentProvider = new WideGlyphProvider(400);
        var childProvider = new WideGlyphProvider();
        var child = new OfficeDrawing(30, 40).AddText("AA", 0, 0, 30, 40,
            new OfficeFontInfo("Nested Font", 20), wrapText: true, shrinkToFit: true);
        child.Fonts.Add("Nested Font", ManagedTextShapingTestAssets.CreateFont('A'));
        var sibling = child.Clone();
        child.ApplyImageExportOptions(new OfficeImageExportOptions { TextShapingProvider = childProvider, TextShapingLanguage = "ar-SA" });
        var parent = new OfficeDrawing(30, 80).ApplyImageExportOptions(new OfficeImageExportOptions {
            TextShapingProvider = parentProvider, TextShapingLanguage = "en-US"
        }).AddClippedDrawing(child, 0, 0, OfficeClipPath.Rectangle(30, 40))
          .AddClippedDrawing(sibling, 0, 40, OfficeClipPath.Rectangle(30, 40));
        var text = XDocument.Parse(OfficeDrawingSvgExporter.ToSvg(parent)).Descendants().Where(element => element.Name.LocalName == "text").ToArray();
        Assert.Equal(new[] { "A", "A", "AA" }, text.Select(element => element.Value));
        Assert.NotEmpty(parentProvider.Languages);
        Assert.All(parentProvider.Languages, value => Assert.Equal("en-US", value));
        Assert.All(childProvider.Languages, value => Assert.Equal("ar-SA", value));
    }

    private static string[] TextElements(string svg) => XDocument.Parse(svg).Descendants()
        .Where(element => element.Name.LocalName == "text").Select(element => element.ToString(SaveOptions.DisableFormatting)).ToArray();

    private sealed class WideGlyphProvider : IOfficeTextShapingProvider {
        private readonly int _advance;
        internal WideGlyphProvider(int advance = 1400) => _advance = advance;
        internal System.Collections.Generic.List<string?> Languages { get; } = new();
        public OfficeTextShapingResult? ShapeText(OfficeTextShapingRequest request) {
            Languages.Add(request.Language);
            return new OfficeTextShapingResult(request.Text.Select((value, index) =>
                new OfficeShapedGlyph(1, value.ToString(), index, advanceWidth: _advance)).ToArray());
        }
    }
}
