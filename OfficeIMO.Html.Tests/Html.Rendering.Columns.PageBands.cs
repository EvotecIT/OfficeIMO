using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using Xunit;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Fact]
    public void HtmlColumns_PageBandsFillBothColumnsBeforeAdvancingThePage() {
        HtmlRenderDocument rendered = RenderColumnPageBands("<section class='page-columns'>"
            + ColumnPageBandTable() + "</section><p id='after'>AFTER</p>");

        Assert.Equal(5, rendered.Pages.Count);
        AssertColumnPageBandTextOnce(rendered, Enumerable.Range(1, 8).Select(index => "LONG-" + index).Append("AFTER"));
        AssertColumnPageBandFourByFour(rendered.Pages[0]);
        Assert.Equal(5, ColumnPageBandTextPage(rendered, "AFTER").PageNumber);
        AssertColumnPageBandNoFragmentLoss(rendered);
    }

    [Fact]
    public void HtmlColumns_PageBandsLeaveTheDirectTableParagraphSplitUnchanged() {
        HtmlRenderDocument rendered = RenderColumnPageBands(ColumnPageBandTable() + "<p id='after'>AFTER</p>");

        AssertColumnPageBandTextOnce(rendered, Enumerable.Range(1, 8).Select(index => "LONG-" + index).Append("AFTER"));
        Assert.Equal(Enumerable.Range(1, 4).Select(index => "LONG-" + index),
            ColumnPageBandText(rendered.Pages[0]).Where(text => text.Text.StartsWith("LONG-", StringComparison.Ordinal)).Select(text => text.Text));
        Assert.Equal(Enumerable.Range(5, 4).Select(index => "LONG-" + index),
            ColumnPageBandText(rendered.Pages[1]).Where(text => text.Text.StartsWith("LONG-", StringComparison.Ordinal)).Select(text => text.Text));
        AssertColumnPageBandNoFragmentLoss(rendered);
    }

    [Fact]
    public void HtmlColumns_PageBandsPreserveOrdinaryParagraphsAndTheirFinalLinkDestination() {
        string paragraphs = string.Concat(Enumerable.Range(1, 24).Select(index =>
            "<p><a href='#after'>FLOW-" + index.ToString("D2") + "</a></p>"));
        string html = ColumnPageBandSource("<section class='page-columns'>" + paragraphs + "</section><p id='after'>AFTER</p>");
        HtmlRenderOptions options = ColumnPageBandOptions();
        HtmlConversionDocument source = HtmlConversionDocument.Parse(html);
        HtmlRenderDocument rendered = HtmlRenderEngine.Execute(source, HtmlRenderRequest.Create(
            HtmlRenderIntentProfile.PrintPaged, HtmlRenderEncoder.DisplayList, options)).Document;

        AssertColumnPageBandTextOnce(rendered, Enumerable.Range(1, 24).Select(index => "FLOW-" + index.ToString("D2")).Append("AFTER"));
        HtmlRenderText leftStart = Assert.Single(ColumnPageBandText(rendered.Pages[0]), text => text.Text == "FLOW-01");
        HtmlRenderText rightStart = Assert.Single(ColumnPageBandText(rendered.Pages[0]), text => text.Text == "FLOW-06");
        Assert.True(leftStart.X < rendered.Pages[0].Width / 2D && rightStart.X >= rendered.Pages[0].Width / 2D);
        Assert.Equal(2, ColumnPageBandTextPage(rendered, "FLOW-11").PageNumber);
        Assert.True(ColumnPageBandTextPage(rendered, "FLOW-24").PageNumber > 1);
        foreach (HtmlRenderPage page in rendered.Pages.Where(page => ColumnPageBandText(page)
                     .Any(text => text.Text.StartsWith("FLOW-", StringComparison.Ordinal)))) {
            foreach (HtmlRenderText text in ColumnPageBandText(page).Where(text => text.Text.StartsWith("FLOW-", StringComparison.Ordinal))) {
                Assert.Equal("#after", text.LinkUri);
                Assert.True(text.X >= page.Margins.Left - 0.001D && text.X + text.Width <= page.Width - page.Margins.Right + 0.001D);
                Assert.True(text.Y + text.Height <= page.Height - page.Margins.Bottom + 0.001D);
            }
        }
        HtmlRenderPage afterPage = ColumnPageBandTextPage(rendered, "AFTER");
        HtmlRenderPage destinationPage = Assert.Single(rendered.Pages, page => page.Visuals
            .OfType<HtmlRenderNamedDestination>().Any(destination => destination.Name == "after"));
        Assert.Same(afterPage, destinationPage);
        AssertColumnPageBandNoFragmentLoss(rendered);

        byte[] bytes = source.RenderToPdfResult(HtmlRenderRequest.Create(
            HtmlRenderIntentProfile.PrintPaged, HtmlRenderEncoder.Pdf, options)).ToBytes();
        PdfCore.PdfReadDocument pdf = PdfCore.PdfReadDocument.Open(bytes);
        var destinations = pdf.NamedDestinations.ToDictionary(destination => destination.Name, destination => destination.PageNumber);
        var links = pdf.Pages.SelectMany(page => page.GetLinkAnnotations()).ToArray();
        Assert.Equal(24, links.Length);
        Assert.All(links, link => {
            int? target = link.DestinationPageNumber;
            if (target == null && link.DestinationName is string name && destinations.TryGetValue(name, out int? pageNumber)) target = pageNumber;
            Assert.Equal(afterPage.PageNumber, target);
        });
    }

    [Fact]
    public void HtmlColumns_PageBandsDeferAnOrphanGroupWhenNeitherPartialColumnFits() {
        HtmlRenderDocument rendered = RenderColumnPageBands("<div style='height:40px'>PRELUDE</div>"
            + "<section class='page-columns'>" + ColumnPageBandTable() + "</section><p id='after'>AFTER</p>");

        AssertColumnPageBandTextOnce(rendered, Enumerable.Range(1, 8).Select(index => "LONG-" + index).Concat(new[] { "PRELUDE", "AFTER" }));
        Assert.DoesNotContain(ColumnPageBandText(rendered.Pages[0]), text => text.Text.StartsWith("LONG-", StringComparison.Ordinal));
        AssertColumnPageBandFourByFour(rendered.Pages[1]);
        AssertColumnPageBandNoFragmentLoss(rendered);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void HtmlColumns_PageBandsMoveAFittingAtomicImageToTheNextAvailableColumn(bool preludeOutside) {
        const string pixel = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNgYAAAAAMAASsJTYQAAAAASUVORK5CYII=";
        const string prelude = "<div style='height:40px'>IMAGE-LEAD</div>";
        string body = (preludeOutside ? prelude : string.Empty) + "<section class='page-columns'>"
            + (preludeOutside ? string.Empty : prelude)
            + "<img id='band-image' style='display:block;width:50px;height:80px' src='data:image/png;base64," + pixel + "'>"
            + "<p>IMAGE-AFTER</p></section><p id='after'>AFTER</p>";
        HtmlRenderDocument rendered = RenderColumnPageBands(body);

        AssertColumnPageBandTextOnce(rendered, new[] { "IMAGE-LEAD", "IMAGE-AFTER", "AFTER" });
        HtmlRenderPage imagePage = Assert.Single(rendered.Pages, page => EnumerateRenderVisuals(page.Scene).OfType<HtmlRenderImage>().Any());
        HtmlRenderImage image = Assert.Single(rendered.Pages.SelectMany(page => EnumerateRenderVisuals(page.Scene)).OfType<HtmlRenderImage>());
        Assert.Equal(preludeOutside ? 2 : 1, imagePage.PageNumber);
        Assert.Equal(50D, image.Width, 3);
        Assert.Equal(80D, image.Height, 3);
        Assert.True(image.Y >= imagePage.Margins.Top - 0.001D && image.Y + image.Height <= imagePage.Height - imagePage.Margins.Bottom + 0.001D);
        Assert.Equal(!preludeOutside, image.X >= imagePage.Width / 2D);
        AssertColumnPageBandNoFragmentLoss(rendered);
    }

    private static string ColumnPageBandTable() =>
        "<table><tbody><tr style='height:150px'><td rowspan='2'><div style='line-height:20px;orphans:4;widows:4'>"
        + string.Join("<br>", Enumerable.Range(1, 8).Select(index => "LONG-" + index))
        + "</div></td></tr><tr style='height:150px'></tr></tbody></table>";

    private static string ColumnPageBandSource(string body) =>
        "<!doctype html><style>@page{size:700px 140px;margin:20px}"
        + "html,body{margin:0;background:white;font:14px/18px PageBandWitness}p{margin:0}"
        + "table{width:300px;table-layout:fixed;border-collapse:separate;border-spacing:0}"
        + "td,th{padding:0;border:0;vertical-align:top;orphans:1;widows:1}"
        + ".page-columns{column-count:2;column-fill:auto;column-gap:60px;height:400px}"
        + "</style>" + body;

    private static HtmlRenderOptions ColumnPageBandOptions() {
        var options = new HtmlRenderOptions {
            ViewportWidth = 700D, ViewportHeight = 140D,
            Margins = HtmlRenderMargins.All(0D), DefaultFontFamily = "PageBandWitness",
            AllowSystemFontFallback = false, ResourceUrlPolicy = HtmlUrlPolicy.CreateEmbeddedResourceProfile()
        };
        string fonts = Path.Combine(RepositoryTestPaths.Find(), "OfficeIMO.TestAssets", "Fonts");
        options.Fonts.Add("PageBandWitness", File.ReadAllBytes(Path.Combine(fonts, "OfficeIMOBaselineSans-Regular.ttf")));
        options.Fonts.Add("PageBandWitness", File.ReadAllBytes(Path.Combine(fonts, "OfficeIMOBaselineSans-Bold.ttf")), OfficeFontStyle.Bold);
        return options;
    }

    private static HtmlRenderDocument RenderColumnPageBands(string body) =>
        HtmlRenderEngine.Execute(HtmlConversionDocument.Parse(ColumnPageBandSource(body)), HtmlRenderRequest.Create(
            HtmlRenderIntentProfile.PrintPaged, HtmlRenderEncoder.DisplayList, ColumnPageBandOptions())).Document;

    private static IEnumerable<HtmlRenderText> ColumnPageBandText(HtmlRenderPage page) => page.Visuals.OfType<HtmlRenderText>();

    private static HtmlRenderPage ColumnPageBandTextPage(HtmlRenderDocument rendered, string label) =>
        Assert.Single(rendered.Pages, page => ColumnPageBandText(page).Any(text => text.Text == label));

    private static void AssertColumnPageBandTextOnce(HtmlRenderDocument rendered, IEnumerable<string> labels) {
        HtmlRenderText[] text = rendered.Pages.SelectMany(ColumnPageBandText).ToArray();
        foreach (string label in labels) Assert.Single(text, item => item.Text == label);
    }

    private static void AssertColumnPageBandFourByFour(HtmlRenderPage page) {
        HtmlRenderText[] text = ColumnPageBandText(page).Where(item => item.Text.StartsWith("LONG-", StringComparison.Ordinal)).ToArray();
        Assert.Equal(Enumerable.Range(1, 4).Select(index => "LONG-" + index),
            text.Where(item => item.X < page.Width / 2D).Select(item => item.Text));
        Assert.Equal(Enumerable.Range(5, 4).Select(index => "LONG-" + index),
            text.Where(item => item.X >= page.Width / 2D).Select(item => item.Text));
        Assert.All(text, item => Assert.True(item.Y + item.Height <= page.Height - page.Margins.Bottom + 0.001D));
    }

    private static void AssertColumnPageBandNoFragmentLoss(HtmlRenderDocument rendered) =>
        Assert.DoesNotContain(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.ForcedFragment
            || diagnostic.Code == HtmlRenderDiagnosticCodes.VisualFragmentUnsupported);
}
