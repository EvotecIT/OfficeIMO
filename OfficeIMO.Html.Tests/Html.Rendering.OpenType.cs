using System.Text;
using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using OfficeIMO.TestAssets;
using PdfCore = OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Fact]
    public void HtmlRendering_OpenTypeCssFlowsThroughLayoutSceneAndSvg() {
        const string text = "OfficeIMO 0123";
        const string html = "<p style=\"margin:0;font-family:'OfficeIMO Shaping Test';"
            + "font-kerning:none;font-variant-ligatures:no-common-ligatures discretionary-ligatures;"
            + "font-variant-numeric:oldstyle-nums tabular-nums slashed-zero;"
            + "font-variant-east-asian:jis04 proportional-width ruby;"
            + "font-feature-settings:'ss01' 2,'dlig' off;font-palette:dark\">" + text + "</p>";
        var provider = new ManagedTextShapingTestAssets.RecordingProvider();
        var options = new HtmlRenderOptions {
            ViewportWidth = 260D,
            ViewportHeight = 80D,
            Margins = HtmlRenderMargins.All(0D),
            TextShapingProvider = provider
        };
        options.Fonts.Add(
            ManagedTextShapingTestAssets.FamilyName,
            ManagedTextShapingTestAssets.CreateFont(text.Distinct().Select(character => (int)character).ToArray()));

        HtmlConversionDocument document = HtmlConversionDocument.Parse(html);
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(document, options);
        HtmlRenderText visual = Assert.Single(
            EnumerateRenderVisuals(rendered.Pages[0].Scene).OfType<HtmlRenderText>(),
            item => item.Text.Contains("OfficeIMO", StringComparison.Ordinal));
        OfficeTextShapingRequest request = Assert.Single(
            provider.Requests,
            item => item.Text.Contains("OfficeIMO", StringComparison.Ordinal));
        string svg = Encoding.UTF8.GetString(document.ExportImage(OfficeImageExportFormat.Svg, options).Bytes);

        Assert.Equal("dark", visual.FontPalette);
        Assert.Equal(0, visual.FeatureSettings.Features["kern"]);
        Assert.Equal(0, visual.FeatureSettings.Features["liga"]);
        Assert.Equal(0, visual.FeatureSettings.Features["clig"]);
        Assert.Equal(0, visual.FeatureSettings.Features["dlig"]);
        Assert.Equal(1, visual.FeatureSettings.Features["onum"]);
        Assert.Equal(1, visual.FeatureSettings.Features["tnum"]);
        Assert.Equal(1, visual.FeatureSettings.Features["zero"]);
        Assert.Equal(1, visual.FeatureSettings.Features["jp04"]);
        Assert.Equal(1, visual.FeatureSettings.Features["pwid"]);
        Assert.Equal(1, visual.FeatureSettings.Features["ruby"]);
        Assert.Equal(2, visual.FeatureSettings.Features["ss01"]);
        Assert.True(request.FeatureSettings.Equals(visual.FeatureSettings));
        Assert.Contains("font-feature-settings=\"&quot;clig&quot; 0", svg, StringComparison.Ordinal);
        Assert.True(HtmlComputedStyleEngine.IsApplicableSupports("(font-feature-settings:'liga' 1)"));
        Assert.True(HtmlComputedStyleEngine.IsApplicableSupports("(font-variant-numeric:tabular-nums)"));
        Assert.True(HtmlComputedStyleEngine.IsApplicableSupports("(font-palette:dark)"));
    }

    [Fact]
    public void HtmlRendering_OpenTypeFeaturesAreInheritedAndCanBeOverridden() {
        const string html = "<div style=\"font-kerning:none;font-feature-settings:'liga' 1\">"
            + "<span id='inherited'>fi</span><span id='overridden' style=\"font-kerning:normal;font-feature-settings:'liga' 0\">fi</span></div>";
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, new HtmlRenderOptions {
            ViewportWidth = 180D,
            ViewportHeight = 80D,
            Margins = HtmlRenderMargins.All(0D)
        });
        HtmlRenderText[] visuals = EnumerateRenderVisuals(rendered.Pages[0].Scene)
            .OfType<HtmlRenderText>()
            .Where(item => item.Text == "fi")
            .ToArray();

        Assert.Equal(2, visuals.Length);
        Assert.Contains(visuals, item => item.FeatureSettings.Features["kern"] == 0 && item.FeatureSettings.Features["liga"] == 1);
        Assert.Contains(visuals, item => item.FeatureSettings.Features["kern"] == 1 && item.FeatureSettings.Features["liga"] == 0);
    }

    [Fact]
    public void HtmlRendering_OpenTypeLigaturesReachRasterAndSearchablePdfOutput() {
        byte[] font = ManagedTextShapingTestAssets.CreateFontWithLigature('f', 'i');
        string fontSource = "@font-face{font-family:'OfficeIMO Shaping Test';src:url('data:font/ttf;base64,"
            + Convert.ToBase64String(font) + "')}";
        string enabledHtml = "<style>" + fontSource
            + "p{margin:0;font:32px 'OfficeIMO Shaping Test';font-feature-settings:'liga' 1}</style><p>fi</p>";
        string disabledHtml = "<style>" + fontSource
            + "p{margin:0;font:32px 'OfficeIMO Shaping Test';font-feature-settings:'liga' 0}</style><p>fi</p>";
        var renderOptions = new HtmlRenderOptions {
            ViewportWidth = 100D,
            ViewportHeight = 60D,
            Margins = HtmlRenderMargins.All(0D)
        };

        byte[] enabledPng = HtmlConversionDocument.Parse(enabledHtml)
            .ExportImage(OfficeImageExportFormat.Png, renderOptions).Bytes;
        byte[] disabledPng = HtmlConversionDocument.Parse(disabledHtml)
            .ExportImage(OfficeImageExportFormat.Png, renderOptions).Bytes;
        var pdfOptions = new HtmlPdfSaveOptions(renderOptions) {
            PdfOptions = new PdfCore.PdfOptions { CompressContentStreams = false }
        };
        byte[] enabledPdf = HtmlConversionDocument.Parse(enabledHtml).ToPdf(pdfOptions);
        byte[] disabledPdf = HtmlConversionDocument.Parse(disabledHtml).ToPdf(pdfOptions);
        int enabledPathCommandCount = PdfCore.PdfDocument.Load(enabledPdf).Render.Drawing(1).Shapes
            .Sum(shape => shape.Shape.PathCommands.Count);
        int disabledPathCommandCount = PdfCore.PdfDocument.Load(disabledPdf).Render.Drawing(1).Shapes
            .Sum(shape => shape.Shape.PathCommands.Count);

        Assert.NotEqual(enabledPng, disabledPng);
        Assert.True(enabledPathCommandCount < disabledPathCommandCount);
        Assert.Equal("fi", PdfCore.PdfReadDocument.Open(enabledPdf).ExtractText().Trim());
        Assert.Equal("fi", PdfCore.PdfReadDocument.Open(disabledPdf).ExtractText().Trim());
    }

    [Fact]
    public void HtmlRendering_FontPaletteSelectsManagedColorGlyphPalette() {
        byte[] font = ManagedTextShapingTestAssets.CreateColorFont('A');
        var options = new HtmlRenderOptions {
            ViewportWidth = 90D,
            ViewportHeight = 60D,
            Margins = HtmlRenderMargins.All(0D)
        };
        options.Fonts.Add("Color Test", font);
        const string shared = "margin:0;font-family:'Color Test';font-size:40px;line-height:48px;color:black;font-palette:";
        HtmlConversionDocument light = HtmlConversionDocument.Parse("<p style=\"" + shared + "light\">A</p>");
        HtmlConversionDocument dark = HtmlConversionDocument.Parse("<p style=\"" + shared + "dark\">A</p>");
        HtmlRenderDocument lightRendered = HtmlRenderTestDriver.Render(light, options);
        HtmlRenderText lightVisual = Assert.Single(lightRendered.Pages[0].Visuals.OfType<HtmlRenderText>());
        OfficeDrawingText lightDrawingText = Assert.Single(lightRendered.Pages[0].CreateDrawing().Elements.OfType<OfficeDrawingText>());

        Assert.Equal("light", lightVisual.FontPalette);
        Assert.Equal("light", lightDrawingText.FontPalette);
        Assert.True(OfficePngReader.TryDecode(light.ExportImage(OfficeImageExportFormat.Png, options).Bytes, out OfficeRasterImage? lightImage));
        Assert.True(OfficePngReader.TryDecode(dark.ExportImage(OfficeImageExportFormat.Png, options).Bytes, out OfficeRasterImage? darkImage));
        string darkSvg = Encoding.UTF8.GetString(dark.ExportImage(OfficeImageExportFormat.Svg, options).Bytes);
        byte[] darkPdf = dark.ToPdf(new HtmlPdfSaveOptions(options));
        OfficeColor[] pdfPaints = PdfCore.PdfDocument.Load(darkPdf).Render.Drawing(1).Shapes
            .Select(shape => shape.Shape.FillColor ?? OfficeColor.Transparent)
            .ToArray();

        Assert.NotNull(lightImage);
        Assert.NotNull(darkImage);
        Assert.True(ContainsColor(lightImage!, pixel => pixel.R > 220 && pixel.G < 40 && pixel.B < 40), DescribeColors(lightImage!));
        Assert.True(ContainsColor(lightImage!, pixel => pixel.B > 220 && pixel.R < 40 && pixel.G < 40));
        Assert.True(ContainsColor(darkImage!, pixel => pixel.R > 220 && pixel.G > 220 && pixel.B < 40));
        Assert.True(ContainsColor(darkImage!, pixel => pixel.G > 90 && pixel.R < 40 && pixel.B < 40));
        Assert.Contains("font-palette=\"dark\"", darkSvg, StringComparison.Ordinal);
        Assert.Contains(OfficeColor.Yellow, pdfPaints);
        Assert.Contains(OfficeColor.FromRgb(0, 128, 0), pdfPaints);
        Assert.Equal("A", PdfCore.PdfReadDocument.Open(darkPdf).ExtractText().Trim());
    }

    private static bool ContainsColor(OfficeRasterImage image, Func<OfficeColor, bool> predicate) {
        for (int y = 0; y < image.Height; y++) {
            for (int x = 0; x < image.Width; x++) {
                if (predicate(image.GetPixel(x, y))) return true;
            }
        }
        return false;
    }

    private static string DescribeColors(OfficeRasterImage image) {
        var counts = new Dictionary<OfficeColor, int>();
        for (int y = 0; y < image.Height; y++) {
            for (int x = 0; x < image.Width; x++) {
                OfficeColor color = image.GetPixel(x, y);
                counts[color] = counts.TryGetValue(color, out int count) ? count + 1 : 1;
            }
        }
        return string.Join(", ", counts.OrderByDescending(pair => pair.Value).Take(12).Select(pair => pair.Key + "=" + pair.Value));
    }
}
