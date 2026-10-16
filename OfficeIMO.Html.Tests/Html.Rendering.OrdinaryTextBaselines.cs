using System.Xml.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Theory]
    [InlineData("1.2", 15D)]
    [InlineData("normal", 15D)]
    [InlineData("20px", 15D)]
    [InlineData("30px", 20D)]
    public void HtmlOrdinaryText_SelectedFaceAndLeadingSupplyTheSavedBaseline(
        string lineHeight, double expectedBaseline) {
        // The byte-backed Baseline Sans native controls have a 15px ascent
        // at 16px. These are layout baselines, before browser PDF snapping.
        string html = "<!doctype html><style>body{margin:0;padding:24px;font:16px/" + lineHeight
            + " Pinned}</style><div>Hgyp</div>";
        HtmlRenderOptions options = OrdinaryBaselineOptions();

        HtmlConversionDocument document = HtmlConversionDocument.Parse(html);
        XElement text = Assert.Single(XDocument.Parse(document.ToSvg(options))
            .Descendants(), node => node.Name.LocalName == "text");

        Assert.Equal(24D + expectedBaseline, HtmlSvgTestCoordinates.CssValue(text, "y"), 3);
        Assert.Equal("Hgyp", text.Value);
        using var pdf = UglyToad.PdfPig.PdfDocument.Open(document.ToPdfBytes(new HtmlToPdfOptions(options)));
        var page = pdf.GetPage(1);
        double savedPdfBaseline = (page.Height - page.Letters.First(letter => letter.Value == "H").StartBaseLine.Y) / 0.75D;
        Assert.Equal(24D + expectedBaseline, savedPdfBaseline, 3);
    }

    [Theory]
    [InlineData("Pinned")]
    [InlineData("Missing, Pinned")]
    public void HtmlOrdinaryText_WrappedAndFallbackRunsKeepSelectedBaselines(string family) {
        string html = "<!doctype html><style>body{margin:0;padding:24px;font:16px/20px Pinned}</style>"
            + "<div style='width:60px;font-family:" + family + "'>Hgyp Hgyp Hgyp Hgyp</div>";
        HtmlRenderOptions options = OrdinaryBaselineOptions();
        XElement[] text = XDocument.Parse(HtmlConversionDocument.Parse(html).ToSvg(options)).Descendants()
            .Where(node => node.Name.LocalName == "text").ToArray();

        Assert.Equal(4, text.Length);
        for (int index = 0; index < text.Length; index++) {
            Assert.Equal(39D + index * 20D, HtmlSvgTestCoordinates.CssValue(text[index], "y"), 3);
            Assert.Equal("Hgyp", text[index].Value.Trim());
        }
    }

    [Fact]
    public void HtmlOrdinaryText_MixedSizesUseTheLargestParticipatingFaceBaseline() {
        const string html = "<!doctype html><style>body{margin:0;padding:24px;font:16px/20px Pinned}</style><div>"
            + "<span style='font-size:24px;line-height:30px'>H</span>"
            + "<span style='font-size:10px;line-height:12px'>gyp</span></div>";
        XElement[] text = XDocument.Parse(HtmlConversionDocument.Parse(html).ToSvg(OrdinaryBaselineOptions()))
            .Descendants().Where(node => node.Name.LocalName == "text").ToArray();

        Assert.Equal(2, text.Length);
        Assert.All(text, node => Assert.Equal(47D, HtmlSvgTestCoordinates.CssValue(node, "y"), 3));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void HtmlOrdinaryText_GridAlignsSelectedBaselinesBeforeRelativePaint(bool relativePaint) {
        string html = "<!doctype html><style>body{margin:0;padding:24px;font:16px/20px Pinned}</style>"
            + "<div style='display:grid;grid-template-columns:100px 100px;grid-template-rows:60px;align-items:baseline'>"
            + "<span style='font-size:24px;line-height:30px'>H</span>"
            + "<span style='font-size:10px;line-height:12px;" + (relativePaint ? "position:relative;top:7px" : "")
            + "'>gyp</span></div>";
        XElement[] text = XDocument.Parse(HtmlConversionDocument.Parse(html).ToSvg(OrdinaryBaselineOptions()))
            .Descendants().Where(node => node.Name.LocalName == "text").ToArray();

        Assert.Equal(2, text.Length);
        Assert.Equal(47D, HtmlSvgTestCoordinates.CssValue(text[0], "y"), 3);
        Assert.Equal(relativePaint ? 54D : 47D,
            HtmlSvgTestCoordinates.CssValue(text[1], "y"), 3);
    }

    [Theory]
    [InlineData("absolute", false)]
    [InlineData("absolute", true)]
    [InlineData("fixed", false)]
    [InlineData("fixed", true)]
    [InlineData("running(header)", false)]
    [InlineData("running(header)", true)]
    public void HtmlOrdinaryText_OutOfFlowOffsetsDoNotChangeFallbackPlacement(string position, bool floated) {
        foreach (string leading in new[] { "0", "1.2" }) {
            HtmlRenderText Render(string marker) {
                string html = "<!doctype html><style>body{margin:0;font:16px/" + leading
                    + " Unavailable}</style>" + (floated ? "<i style='float:left;width:20px;height:160px'></i>" : "")
                    + "<div>BODY" + marker + "</div>";
                var options = new HtmlRenderOptions {
                    ViewportWidth = 400D, Margins = HtmlRenderMargins.All(0D),
                    UserAgentStyles = HtmlRenderUserAgentStyleMode.Browser, AllowSystemFontFallback = false
                };
                HtmlRenderDocument rendered = HtmlRenderEngine.Execute(HtmlConversionDocument.Parse(html),
                    HtmlRenderRequest.Create(HtmlRenderIntentProfile.ScreenFullPage, options: options)).Document;
                return Assert.Single(EnumerateRenderVisuals(rendered.Pages[0].Scene).OfType<HtmlRenderText>(),
                    text => text.Text == "BODY");
            }

            HtmlRenderText control = Render(string.Empty);
            HtmlRenderText actual = Render("<span style='position:" + position
                + ";vertical-align:20px;left:200px;top:200px'>OUTSIDE</span>");
            Assert.Equal(control.Y, actual.Y, 6);
            Assert.Equal(control.Height, actual.Height, 6);
            Assert.Equal(control.LineHeight, actual.LineHeight, 6);
        }
    }

    private static HtmlRenderOptions OrdinaryBaselineOptions() {
        var options = new HtmlRenderOptions {
            ViewportWidth = 400D, ViewportHeight = 200D,
            Margins = HtmlRenderMargins.All(0D),
            UserAgentStyles = HtmlRenderUserAgentStyleMode.Browser
        };
        options.Fonts.Add("Pinned", File.ReadAllBytes(Path.Combine(RepositoryTestPaths.Find(),
            "OfficeIMO.TestAssets", "Fonts", "OfficeIMOBaselineSans-Regular.ttf")));
        return options;
    }

    [Fact]
    public void HtmlOrdinaryText_NativePdfGlyphsDoNotReplaceTheQualifiedImageLineCarrier() {
        string image = "<img style='width:1px;height:0' src='data:image/png;base64,"
            + Convert.ToBase64String(OfficePngWriter.Encode(new OfficeRasterImage(1, 1, OfficeColor.Red))) + "'>";
        OfficeIMO.Pdf.PdfReadDocument Read(string sibling) => OfficeIMO.Pdf.PdfReadDocument.Open(
            HtmlConversionDocument.Parse("<style>body{margin:0;font:16px/20px Pinned}</style><div>H"
                + sibling + "</div>").ToPdfBytes(new HtmlToPdfOptions(OrdinaryBaselineOptions())));
        OfficeIMO.Pdf.PdfReadDocument ordinary = Read(string.Empty), imageLine = Read(image);

        Assert.True(Assert.Single(ordinary.Pages[0].GetTextSpans(), span => span.Text == "H").IsVisible);
        Assert.False(Assert.Single(imageLine.Pages[0].GetTextSpans(), span => span.Text == "H").IsVisible);
        Assert.Equal("H", ordinary.ExtractText().Trim());
        Assert.Equal("H", imageLine.ExtractText().Trim());
    }
}
