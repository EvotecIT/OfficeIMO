using System.IO;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using Xunit;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Theory]
    [InlineData(HtmlRenderUserAgentStyleMode.Document)]
    [InlineData(HtmlRenderUserAgentStyleMode.Browser)]
    public void HtmlRendering_CodeKeepsUserAgentMonospaceAgainstInheritedFamily(HtmlRenderUserAgentStyleMode mode) {
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(
            "<article style='font-family:Georgia'><p>Body <code>inline</code></p><pre>block</pre><p style='font-family:Arial'><code style='font-family:inherit'>inherited</code></p></article>",
            new HtmlRenderOptions { UserAgentStyles = mode });
        HtmlRenderText[] texts = rendered.Pages.SelectMany(page => EnumerateMathMlScene(page.Scene)).OfType<HtmlRenderText>().ToArray();

        Assert.DoesNotContain("Georgia", texts.First(text => text.Text.Trim() == "inline").Font.FamilyName);
        Assert.DoesNotContain("Georgia", texts.First(text => text.Text.Trim() == "block").Font.FamilyName);
        Assert.Contains("Georgia", texts.First(text => text.Text.Trim() == "Body").Font.FamilyName);
        // An explicit inherit still inherits.
        Assert.Contains("Arial", texts.First(text => text.Text.Trim() == "inherited").Font.FamilyName);
    }

    [Fact]
    public void HtmlRendering_SmallCapsDrawLowercaseAsSmallerCapitals() {
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render("<p style='font-size:20px;font-variant:small-caps'>Report</p>");
        HtmlRenderText[] texts = rendered.Pages.SelectMany(page => EnumerateMathMlScene(page.Scene)).OfType<HtmlRenderText>()
            .Where(text => text.Text.Trim().Length > 0).ToArray();

        HtmlRenderText capital = Assert.Single(texts, text => text.Text == "R");
        HtmlRenderText small = Assert.Single(texts, text => text.Text == "EPORT");
        Assert.True(small.Font.Size < capital.Font.Size * 0.8D, $"{small.Font.Size} vs {capital.Font.Size}");
        Assert.True(small.Font.Size > capital.Font.Size * 0.6D, $"{small.Font.Size} vs {capital.Font.Size}");
    }

    [Fact]
    public void HtmlRendering_UnloadedImageWithoutSizeTakesItsAlternativeTextSpace() {
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(
            "<p><img src='https://example.invalid/badge.svg' alt='Build'> · Documentation</p>");
        HtmlRenderVisual[] scene = rendered.Pages.SelectMany(page => EnumerateMathMlScene(page.Scene)).ToArray();
        HtmlRenderText documentation = Assert.Single(scene.OfType<HtmlRenderText>(), text => text.Text.Contains("Documentation"));
        HtmlRenderText alternative = Assert.Single(scene.OfType<HtmlRenderText>(), text => text.Text == "Build");

        // The following text stays on the same line instead of being pushed below a 300x150 box.
        Assert.True(Math.Abs(documentation.Y - alternative.Y) < alternative.Height, $"{documentation.Y} vs {alternative.Y}");
        Assert.True(documentation.X < 150D, documentation.X.ToString(System.Globalization.CultureInfo.InvariantCulture));
    }

    [Fact]
    public void HtmlRendering_SizedUnloadedImageKeepsItsAuthoredBox() {
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(
            "<p><img src='https://example.invalid/badge.svg' alt='Build' width='120' height='40'>After</p>");
        HtmlRenderText after = Assert.Single(rendered.Pages.SelectMany(page => EnumerateMathMlScene(page.Scene)).OfType<HtmlRenderText>(),
            text => text.Text.Contains("After"));
        Assert.True(after.X >= 120D, after.X.ToString(System.Globalization.CultureInfo.InvariantCulture));
    }

    [Fact]
    public void HtmlToPdf_TrustedHostLoadsLocalFileImages() {
        string directory = Path.Combine(Path.GetTempPath(), "officeimo-html-local-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        try {
            string image = Path.Combine(directory, "pixel.png");
            File.WriteAllBytes(image, Convert.FromBase64String(
                "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR4nGP4z8DwHwAFAAH/iZk9HQAAAABJRU5ErkJggg=="));
            string htmlPath = Path.Combine(directory, "page.html");
            File.WriteAllText(htmlPath, "<p><img src='pixel.png' alt='Pixel' width='20' height='20'></p>");
            HtmlConversionDocument document = HtmlConversionDocument.Load(htmlPath, new HtmlConversionDocumentOptions {
                Trust = HtmlInputTrust.Trusted,
                BaseUri = new Uri(directory + Path.DirectorySeparatorChar)
            });

            var trusted = new HtmlToPdfOptions { ResourcePolicy = PdfCore.PdfResourcePolicy.CreateTrustedHost() };
            byte[] pdf = document.ToPdf(trusted).ToBytes();
            Assert.Single(PdfCore.PdfImageExtractor.ExtractImages(pdf));

            byte[] untrusted = document.ToPdf(new HtmlToPdfOptions()).ToBytes();
            Assert.Empty(PdfCore.PdfImageExtractor.ExtractImages(untrusted));
        } finally {
            Directory.Delete(directory, recursive: true);
        }
    }

    [Fact]
    public void HtmlRendering_AbsolutelyPositionedCheckboxKeepsItsControlSize() {
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(
            "<ul style='list-style:none;padding-left:0'><li style='position:relative;padding-left:1.6rem'>" +
            "<input type='checkbox' checked style='position:absolute;left:0;top:.2em'> Signed</li></ul>");
        HtmlRenderFormField box = Assert.Single(rendered.Pages.SelectMany(page => EnumerateMathMlScene(page.Scene)).OfType<HtmlRenderFormField>());
        Assert.True(box.Width >= 12D, box.Width.ToString(System.Globalization.CultureInfo.InvariantCulture));
        Assert.True(box.Height >= 12D, box.Height.ToString(System.Globalization.CultureInfo.InvariantCulture));
    }

    [Fact]
    public void HtmlRendering_TrailingNewlineInPreAddsNoEmptyLine() {
        HtmlRenderDocument one = HtmlRenderTestDriver.Render("<pre style='margin:0;line-height:20px'><code>a\nb</code></pre><p style='margin:0'>after</p>");
        HtmlRenderDocument trailing = HtmlRenderTestDriver.Render("<pre style='margin:0;line-height:20px'><code>a\nb\n</code></pre><p style='margin:0'>after</p>");
        double Y(HtmlRenderDocument document) => document.Pages.SelectMany(page => EnumerateMathMlScene(page.Scene))
            .OfType<HtmlRenderText>().Single(text => text.Text == "after").Y;

        Assert.Equal(Y(one), Y(trailing), 1);
    }
}
