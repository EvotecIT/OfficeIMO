using System.Globalization;
using System.Xml.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Theory]
    [InlineData("16px", "1.2", 74.1275D)]
    [InlineData("16px", "19.2px", 74.143125D)]
    [InlineData("16px", "20px", 74.94D)]
    [InlineData("16px", "21px", 74.94D)]
    [InlineData("16px", "18px", 73.94D)]
    [InlineData("20px", "1.2", 74.94D)]
    public void HtmlImageLine_SelectedFaceAndAuthoredLeadingPlaceFollowingBlock(
        string fontSize, string lineHeight, double followingTop) {
        // Native Chromium witnesses use the same licensed Baseline Sans face.
        // These positions retain the renderer's authored 53.94px image height;
        // the browser's image is 53.9375px on its CSS coordinate grid.
        string html = ImageLineSource("<div style='font-size:" + fontSize + ";line-height:" + lineHeight + "'>"
            + NormalLineImage + "</div>" + NormalLineFollowing);
        HtmlRenderDocument rendered = RenderImageLine(html, ImageLineOptions());

        Assert.Equal(53.94D, NormalLineImageVisual(rendered).Height, 6);
        Assert.Equal(followingTop, NormalLineFollowingVisual(rendered).Y, 6);
    }

    [Theory]
    [InlineData("1.2", 72.94D)]
    [InlineData("19.2px", 76.143125D)]
    public void HtmlImageLine_InheritedMultiplierAndLengthKeepTheirOwnSizeBasis(
        string lineHeight, double followingTop) {
        string html = ImageLineSource("<div style='font-size:16px;line-height:" + lineHeight + "'>"
            + "<div style='font-size:20px'><div style='font-size:10px'>" + NormalLineImage
            + "</div>" + NormalLineFollowing + "</div></div>");
        HtmlRenderDocument rendered = RenderImageLine(html, ImageLineOptions());

        Assert.Equal(followingTop, NormalLineFollowingVisual(rendered).Y, 6);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void HtmlImageLine_TextAndWrappedFloatFlowShareTheSavedBaseline(bool floated) {
        string floatBox = floated ? "<i style='float:left;width:20px;height:160px'></i>" : string.Empty;
        string html = ImageLineSource(floatBox + "<div style='width:200px'>" + NormalLineImage
            + "<span>Hgyp</span><br>TAIL</div>" + NormalLineFollowing);
        HtmlConversionDocument document = HtmlConversionDocument.Parse(html);
        HtmlRenderOptions options = ImageLineOptions();
        HtmlRenderDocument rendered = HtmlRenderEngine.Execute(document,
            HtmlRenderRequest.Create(HtmlRenderIntentProfile.PrintPaged, options: options)).Document;
        HtmlRenderImage image = NormalLineImageVisual(rendered);
        HtmlRenderText tail = Assert.Single(EnumerateRenderVisuals(rendered.Pages[0].Scene)
            .OfType<HtmlRenderText>(), item => item.Text == "TAIL");

        // The following ordinary line now has its selected-font paint anchor;
        // its flow origin and the image line's saved baseline remain separate.
        Assert.Equal(74.1275D, tail.LayoutY, 6);
        Assert.Equal(93.3275D, NormalLineFollowingVisual(rendered).Y, 6);

        XElement savedText = Assert.Single(XDocument.Parse(document.ToSvg(options)).Descendants()
            .Where(node => node.Name.LocalName == "text"), node => node.Value == "Hgyp");
        double savedBaseline = double.Parse(savedText.Attribute("y")!.Value, CultureInfo.InvariantCulture);
        Assert.Equal(image.Y + image.Height + 16D, savedBaseline, 3);

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(document.ToPdfBytes(new HtmlToPdfOptions(options)));
        var page = pdf.GetPage(1);
        var first = page.Letters.First(letter => letter.Value == "H");
        double pdfBaselineCss = (page.Height - first.StartBaseLine.Y) / 0.75D;
        Assert.Equal(savedBaseline, pdfBaselineCss, 3);
    }

    private static readonly string NormalLineImage = "<img id='picture' style='width:64px;height:53.94px;margin-bottom:16px' "
        + "src='data:image/png;base64," + Convert.ToBase64String(OfficeIMO.Tests.Pdf.PdfPngTestImages.CreateRgbPng(255, 0, 0)) + "'>";
    private const string NormalLineFollowing = "<div id='following' style='height:2px;background:blue'></div>";

    private static string ImageLineSource(string content) => "<!doctype html><style>"
        + "html{font-size:16px}body{margin:0;padding:0;font:16px/1.2 Pinned}</style>" + content;

    private static HtmlRenderOptions ImageLineOptions() {
        var options = new HtmlRenderOptions {
            Mode = HtmlRenderMode.Paged,
            ViewportWidth = 816D, ViewportHeight = 900D, Margins = HtmlRenderMargins.All(0D),
            PageSize = new OfficePageSize(816D / 96D, 900D / 96D),
            UserAgentStyles = HtmlRenderUserAgentStyleMode.Browser, HonorCssPageRules = false
        };
        options.Fonts.Add("Pinned", File.ReadAllBytes(Path.Combine(RepositoryTestPaths.Find(),
            "OfficeIMO.TestAssets", "Fonts", "OfficeIMOBaselineSans-Regular.ttf")));
        return options;
    }

    private static HtmlRenderDocument RenderImageLine(string html, HtmlRenderOptions options) =>
        HtmlRenderEngine.Execute(HtmlConversionDocument.Parse(html),
            HtmlRenderRequest.Create(HtmlRenderIntentProfile.PrintPaged, options: options)).Document;

    private static HtmlRenderImage NormalLineImageVisual(HtmlRenderDocument document) => Assert.Single(
        document.Pages.SelectMany(page => EnumerateRenderVisuals(page.Scene)).OfType<HtmlRenderImage>(),
        image => image.Source == "img#picture");

    private static HtmlRenderShape NormalLineFollowingVisual(HtmlRenderDocument document) => Assert.Single(
        document.Pages.SelectMany(page => EnumerateRenderVisuals(page.Scene)).OfType<HtmlRenderShape>(),
        shape => shape.Source == "div#following" && shape.Shape.FillColor.HasValue);
}
