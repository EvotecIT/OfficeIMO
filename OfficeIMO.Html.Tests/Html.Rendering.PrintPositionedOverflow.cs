using System.Text;
using System.Threading.Tasks;
using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using PdfCore = OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Theory]
    [InlineData("ratio", 1, 4)]
    [InlineData("ratio", 9, 3)]
    [InlineData("absolute", 1, 4)]
    [InlineData("absolute", 9, 3)]
    [InlineData("fixed", 1, 4)]
    [InlineData("fixed", 9, 4)]
    public async Task PrintFitCountsAbsoluteCardOverflowWithoutFittingInFlowHeaders(
        string header, int cardCount, int expectedPages) {
        string html = CreatePositionedPrintOverflowFixture(header, cardCount);
        var options = new HtmlToPdfOptions {
            PageSize = OfficePageSizes.A4,
            Margins = HtmlRenderMargins.All(0D),
            ViewportWidth = 816D
        };

        var result = await HtmlConversionDocument.Parse(html).RenderToPdfResultAsync(
            HtmlRenderRequest.Create(HtmlRenderIntentProfile.PrintPaged, HtmlRenderEncoder.Pdf, options));

        Assert.Equal(expectedPages, result.RenderResult.Document.Pages.Count);
        Assert.DoesNotContain(result.RenderResult.Document.Diagnostics,
            diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.PrintFitOverflowUnresolved);
        string pdfText = PdfCore.PdfReadDocument.Open(result.ToBytes()).ExtractText();
        for (int number = 1; number <= 30; number++) {
            Assert.Contains("PARAGRAPH " + number, pdfText);
        }
    }

    [Fact]
    public async Task ExplicitlyDisabledPrintFitKeepsAbsoluteCardOverflowAtFullScale() {
        var options = new HtmlToPdfOptions {
            PageSize = OfficePageSizes.A4,
            Margins = HtmlRenderMargins.All(0D),
            ViewportWidth = 816D,
            AutoFitWidePrintContent = false
        };
        var result = await HtmlConversionDocument.Parse(CreatePositionedPrintOverflowFixture("absolute", 9))
            .RenderToPdfResultAsync(HtmlRenderRequest.Create(
                HtmlRenderIntentProfile.PrintPaged, HtmlRenderEncoder.Pdf, options));

        Assert.Equal(4, result.RenderResult.Document.Pages.Count);
    }

    [Theory]
    [InlineData("transform:translateX(0)")]
    [InlineData("opacity:.9")]
    [InlineData("clip-path:inset(0)")]
    public async Task PrintFitFindsAbsoluteCardBoxInsidePaintEffects(string effect) {
        string html = CreatePositionedPrintOverflowFixture("absolute", 9)
            .Replace(".head{position:absolute;width:100%;height:180px;background:#acf}",
                ".head{position:absolute;width:100%;height:180px;background:#acf;" + effect + "}",
                StringComparison.Ordinal);
        var options = new HtmlToPdfOptions {
            PageSize = OfficePageSizes.A4,
            Margins = HtmlRenderMargins.All(0D),
            ViewportWidth = 816D
        };

        var result = await HtmlConversionDocument.Parse(html).RenderToPdfResultAsync(
            HtmlRenderRequest.Create(HtmlRenderIntentProfile.PrintPaged, HtmlRenderEncoder.Pdf, options));

        Assert.Equal(3, result.RenderResult.Document.Pages.Count);
        Assert.DoesNotContain(result.RenderResult.Document.Diagnostics,
            diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.PrintFitOverflowUnresolved);
    }

    [Fact]
    public async Task InFlowPrintFitDoesNotRefitGrowingPercentagePositionedBox() {
        string html = CreateGrowingPercentagePositionedBoxFixture(110, rootMinWidth: false);
        var options = new HtmlToPdfOptions {
            PageSize = OfficePageSizes.A4,
            Margins = HtmlRenderMargins.All(0D),
            ViewportWidth = 816D
        };

        var result = await HtmlConversionDocument.Parse(html).RenderToPdfResultAsync(
            HtmlRenderRequest.Create(HtmlRenderIntentProfile.PrintPaged, HtmlRenderEncoder.Pdf, options));

        Assert.Equal(3, result.RenderResult.Document.Pages.Count);
        Assert.DoesNotContain(result.RenderResult.Document.Diagnostics,
            diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.PrintFitOverflowUnresolved);
        string pdfText = PdfCore.PdfReadDocument.Open(result.ToBytes()).ExtractText();
        for (int number = 1; number <= 30; number++) Assert.Contains("PARAGRAPH " + number, pdfText);
    }

    [Theory]
    [InlineData(110)]
    [InlineData(140)]
    public async Task RootMinWidthPrintFitDoesNotRefitGrowingPercentagePositionedBox(int percentage) {
        var options = new HtmlToPdfOptions {
            PageSize = OfficePageSizes.A4,
            Margins = HtmlRenderMargins.All(0D),
            ViewportWidth = 816D
        };
        var result = await HtmlConversionDocument.Parse(
                CreateGrowingPercentagePositionedBoxFixture(percentage, rootMinWidth: true))
            .RenderToPdfResultAsync(HtmlRenderRequest.Create(
                HtmlRenderIntentProfile.PrintPaged, HtmlRenderEncoder.Pdf, options));

        Assert.Equal(3, result.RenderResult.Document.Pages.Count);
        Assert.DoesNotContain(result.RenderResult.Document.Diagnostics,
            diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.PrintFitOverflowUnresolved);
    }

    private static string CreateGrowingPercentagePositionedBoxFixture(int percentage, bool rootMinWidth) =>
        "<!doctype html><style>html,body{margin:0;font:16px Arial}"
        + "body{display:flex;flex-direction:column;overflow-x:hidden;"
        + (rootMinWidth ? "min-width:960px" : string.Empty) + "}"
        + (rootMinWidth ? string.Empty : "#wide{width:960px;height:10px;flex:none}")
        + "#relative{position:relative;width:100%;height:100px;overflow:auto;flex:none}"
        + "#absolute{position:absolute;top:0;left:0;width:" + percentage + "%;height:100px;background:#acf}"
        + "aside div{height:100px}p{font:20px Arial;height:30px;margin:0}</style>"
        + "<p>STARTMARK</p>"
        + (rootMinWidth ? string.Empty : "<div id='wide'>WIDE</div>")
        + "<div id='relative'><div id='absolute'>ABSOLUTE</div></div><aside>"
        + string.Concat(Enumerable.Range(1, 30).Select(number => "<div>PARAGRAPH " + number + "</div>"))
        + "</aside>";

    private static string CreatePositionedPrintOverflowFixture(string header, int cardCount) {
        string headerCss = header switch {
            "ratio" => ".head{position:relative;width:100%}.ratio{position:relative;width:100%}"
                + ".ratio::before{display:block;padding-top:56.25%;content:''}"
                + ".ratio>*{position:absolute;top:0;left:0;width:100%;height:100%;background:#acf}",
            "absolute" => ".head{position:absolute;width:100%;height:180px;background:#acf}",
            _ => ".head{height:180px;background:#acf}"
        };
        string cardHeader = header == "ratio"
            ? "<div class='head'><div class='ratio'><div>Image</div></div></div>"
            : "<div class='head'>Header</div>";
        var html = new StringBuilder("<!doctype html><style>html,body{margin:0;font:16px Arial}"
            + "body{display:flex;flex-direction:column;min-height:100vh;overflow-x:hidden}"
            + "ul{display:flex;flex-wrap:nowrap;overflow-x:auto;gap:8px;list-style:none;margin:0 0 16px;padding:0}"
            + "li{display:flex;flex-direction:column;flex:none;width:320px;max-width:80vw;height:450px;"
            + "position:relative;overflow:hidden;background:#ddd}"
            + "aside{flex:none}aside div{height:100px}p{font:20px Arial;height:30px;flex:none;margin:0}</style><style>")
            .Append(headerCss).Append("</style><p>STARTMARK</p><ul>");
        for (int number = 1; number <= cardCount; number++) {
            html.Append("<li>").Append(cardHeader).Append("Card ").Append(number).Append("</li>");
        }
        html.Append("</ul><aside>");
        for (int number = 1; number <= 30; number++) {
            html.Append("<div>PARAGRAPH ").Append(number).Append("</div>");
        }
        return html.Append("</aside>").ToString();
    }
}
