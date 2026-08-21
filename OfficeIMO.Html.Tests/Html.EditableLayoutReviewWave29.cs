using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.Word;
using OfficeIMO.Word.Html;
using System.Text;
using System.Threading.Tasks;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class HtmlEditableLayoutReviewWave29Tests {
    [Fact]
    public void ElementSelectorStylesheetBoundaryIsDiagnosed() {
        string path = Path.GetTempFileName();
        try {
            File.WriteAllText(path,
                "div{position:absolute;left:24px;top:18px;width:160px;height:40px}");
            var options = new HtmlToWordOptions();
            options.StylesheetPaths.Add(path);

            HtmlToWordResult result = HtmlConversionDocument.Parse(
                "<div>Element styled anchor</div>").ToWordDocumentResult(options);
            using WordDocument word = result.Value;

            Assert.Empty(word.TextBoxes);
            Assert.NotEmpty(word.Find("Element styled anchor", StringComparison.Ordinal));
            Assert.Contains(result.Report.Diagnostics, diagnostic =>
                diagnostic.Code == HtmlEditableLayoutDiagnosticCodes.PlacementSimplified
                && diagnostic.Detail == "externalStylesheetSources=true; semanticFlow=true");
        } finally {
            File.Delete(path);
        }
    }

    [Theory]
    [InlineData("<small>Small</small>")]
    [InlineData("<big>Big</big>")]
    [InlineData("<font size='5'>Font</font>")]
    [InlineData("<tt>Teletype</tt>")]
    [InlineData("<strike>Strike</strike>")]
    [InlineData("<center>Centered</center>")]
    public void SemanticFormattingElementsStayInDiagnosedFlow(string content) {
        string html = "<div style='position:absolute;width:220px;height:40px'>" + content + "</div>";

        HtmlEditableLayoutProjection projection = HtmlEditableLayoutProjector.Project(
            HtmlConversionDocument.Parse(html));

        Assert.Empty(projection.Regions);
        Assert.Contains(projection.Diagnostics, diagnostic =>
            diagnostic.Code == HtmlEditableLayoutDiagnosticCodes.PlacementSimplified
            && diagnostic.Detail == "semanticContent=true");
    }

    [Theory]
    [InlineData(OfficePageOrientation.Portrait)]
    [InlineData(OfficePageOrientation.Landscape)]
    public void ContinuousWordProjectionUsesConfiguredPageWidth(OfficePageOrientation orientation) {
        const string html = "<div style='position:absolute;right:0;top:0;width:96px;height:40px'>A5 anchor</div>";
        var options = new HtmlToWordOptions {
            DefaultPageSize = WordPageSize.A5,
            DefaultOrientation = orientation
        };

        HtmlToWordResult result = HtmlConversionDocument.Parse(html).ToWordDocumentResult(options);
        using WordDocument word = result.Value;

        double widthTwips = orientation == OfficePageOrientation.Landscape
            ? WordPageSizes.A5.HeightTwips
            : WordPageSizes.A5.WidthTwips;
        double pageWidth = widthTwips / 1440D * HtmlRenderOptions.CssPixelsPerInch;
        int expectedOffset = (int)Math.Round((pageWidth - 48D - 96D) * 9525D);
        Assert.Equal(expectedOffset, Assert.Single(word.TextBoxes).HorizontalPositionOffset);
    }

    [Fact]
    public async Task RedirectedStylesheetAliasesShareBudgetState() {
        const string css = ".target{color:red}";
        int cssBytes = Encoding.UTF8.GetByteCount(css);
        var limits = HtmlConversionLimits.CreateUntrustedProfile();
        limits.MaxCssBytes = cssBytes;
        limits.MaxTotalCssBytes = cssBytes;
        HtmlConversionDocument document = HtmlConversionDocument.Parse(
            "<link rel='stylesheet' href='https://assets.example.test/a.css'>" +
            "<link rel='stylesheet' href='https://assets.example.test/b.css'>" +
            "<p class='target'>Aliased stylesheet</p>",
            new HtmlConversionDocumentOptions { Limits = limits });
        var options = new HtmlRenderOptions {
            ResourceResolver = (request, cancellationToken) =>
                Task.FromResult<HtmlResolvedResource?>(new HtmlResolvedResource(
                    Encoding.UTF8.GetBytes(css),
                    "text/css",
                    new Uri("https://cdn.example.test/shared.css")))
        };

        HtmlRenderDocument rendered = await HtmlRenderTestDriver.RenderAsync(document, options);

        Assert.DoesNotContain(rendered.Diagnostics, diagnostic =>
            diagnostic.Code == HtmlConversionDiagnosticCodes.CssTotalSizeLimitExceeded);
        Assert.Contains(rendered.Pages[0].Visuals.OfType<HtmlRenderText>(), text =>
            text.Text.Contains("Aliased stylesheet", StringComparison.Ordinal)
            && text.Color == OfficeColor.FromRgb(255, 0, 0));
    }

    [Fact]
    public async Task RedirectedStylesheetAliasesShareRejectedBudgetState() {
        const string seedCss = ".seed{color:blue}";
        const string externalCss = ".target{color:red}";
        int seedBytes = Encoding.UTF8.GetByteCount(seedCss);
        int externalBytes = Encoding.UTF8.GetByteCount(externalCss);
        var limits = HtmlConversionLimits.CreateUntrustedProfile();
        limits.MaxCssBytes = Math.Max(seedBytes, externalBytes);
        limits.MaxTotalCssBytes = seedBytes + externalBytes - 1;
        HtmlConversionDocument document = HtmlConversionDocument.Parse(
            "<style>" + seedCss + "</style>" +
            "<link rel='stylesheet' href='https://assets.example.test/a.css'>" +
            "<link rel='stylesheet' href='https://assets.example.test/b.css'>" +
            "<p class='target'>Aliased stylesheet</p>",
            new HtmlConversionDocumentOptions { Limits = limits });
        var options = new HtmlRenderOptions {
            ResourceResolver = (request, cancellationToken) =>
                Task.FromResult<HtmlResolvedResource?>(new HtmlResolvedResource(
                    Encoding.UTF8.GetBytes(externalCss),
                    "text/css",
                    new Uri("https://cdn.example.test/shared.css")))
        };

        HtmlRenderDocument rendered = await HtmlRenderTestDriver.RenderAsync(document, options);

        Assert.Single(rendered.Diagnostics, diagnostic =>
            diagnostic.Code == HtmlConversionDiagnosticCodes.CssTotalSizeLimitExceeded);
    }
}