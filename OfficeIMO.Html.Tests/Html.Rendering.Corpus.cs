using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using PdfCore = OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    public static IEnumerable<object[]> HtmlRenderingCorpusScenarioIds => HtmlRenderingCorpus.All
        .Select(item => new object[] { item.Id });

    [Fact]
    public void HtmlRenderingCorpus_CoversEveryPublishedMarketScenario() {
        Assert.Equal(
            HtmlMarketScenarioCatalog.All.Select(item => item.Id),
            HtmlRenderingCorpus.All.Select(item => item.Id));
    }

    [Fact]
    public void HtmlRenderingCorpus_DashboardHeadingAndIncidentRemainFullyVisible() {
        HtmlRenderingCorpusCase scenario = HtmlRenderingCorpus.All.Single(item => item.Id == "dashboard-print");
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(scenario.Html, scenario.CreateOptions());
        HtmlRenderText[] text = rendered.Pages.SelectMany(page => EnumerateCorpusVisuals(page.Scene))
            .OfType<HtmlRenderText>()
            .ToArray();
        HtmlRenderText[] heading = text
            .Where(fragment => fragment.Text.Contains("Documents", StringComparison.Ordinal)
                || fragment.Text.Contains("processed", StringComparison.Ordinal))
            .ToArray();
        HtmlRenderText[] incident = text
            .Where(fragment => fragment.Text.Contains("Open incident", StringComparison.Ordinal)
                || fragment.Text.Contains("remapping", StringComparison.Ordinal))
            .ToArray();

        Assert.NotEmpty(heading);
        Assert.NotEmpty(incident);
        Assert.Single(heading.Select(fragment => Math.Round(fragment.Y, 3)).Distinct());
        Assert.Single(incident.Select(fragment => Math.Round(fragment.Y, 3)).Distinct());
    }

    [Theory]
    [MemberData(nameof(HtmlRenderingCorpusScenarioIds))]
    public void HtmlRenderingCorpus_ProvesSharedSceneImageAndSearchablePdf(string scenarioId) {
        HtmlRenderingCorpusCase scenario = HtmlRenderingCorpus.All.Single(item => item.Id == scenarioId);
        HtmlRenderOptions options = scenario.CreateOptions();

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(scenario.Html, options);

        Assert.Equal(scenario.Mode, rendered.Mode);
        Assert.Equal(scenario.ExpectedPageCount, rendered.Pages.Count);
        Assert.All(rendered.Pages, page => {
            Assert.Equal(scenario.ExpectedSurfaceWidth, page.Width, 3);
            Assert.True(page.Height > 0D);
            Assert.True(
                page.Visuals.Count >= scenario.MinimumVisualCount,
                scenario.Id + " page " + page.PageNumber + " produced " + page.Visuals.Count + " visuals; expected at least " + scenario.MinimumVisualCount + ".");
        });
        Assert.True(rendered.Headings.Count >= scenario.MinimumHeadingCount);
        string logicalText = NormalizeCorpusWhitespace(rendered.Text);
        foreach (string marker in scenario.TextMarkers) Assert.Contains(NormalizeCorpusWhitespace(marker), logicalText, StringComparison.Ordinal);
        foreach (string code in scenario.DiagnosticCodes) {
            Assert.Contains(rendered.Diagnostics, diagnostic => diagnostic.Code == code);
        }
        foreach (string code in scenario.ForbiddenDiagnosticCodes) {
            Assert.DoesNotContain(rendered.Diagnostics, diagnostic => diagnostic.Code == code);
        }
        Assert.DoesNotContain(rendered.Diagnostics, diagnostic => diagnostic.Severity == HtmlDiagnosticSeverity.Error);
        HtmlRenderVisual[] visuals = rendered.Pages.SelectMany(page => EnumerateCorpusVisuals(page.Scene)).ToArray();
        foreach (string source in scenario.RequiredVisualSources) {
            Assert.Contains(visuals, visual => string.Equals(visual.Source, source, StringComparison.Ordinal));
        }
        if (scenario.LinkUri != null) {
            Assert.Contains(visuals, visual => visual.LinkUri == scenario.LinkUri);
        }

        OfficeDrawing firstPage = rendered.Pages[0].CreateDrawing();
        byte[] png = OfficeDrawingRasterRenderer.ToPng(firstPage, 0.5D, OfficeColor.White);
        string svg = OfficeDrawingSvgExporter.ToSvg(firstPage, 0.5D);
        Assert.True(png.Length > 100);
        Assert.Equal(new byte[] { 137, 80, 78, 71, 13, 10, 26, 10 }, png.Take(8).ToArray());
        Assert.Contains("<svg", svg, StringComparison.Ordinal);
        foreach (string word in NormalizeCorpusWhitespace(scenario.TextMarkers[0]).Split(' ')) {
            Assert.Contains(word, svg, StringComparison.Ordinal);
        }

        HtmlPdfSaveOptions pdfOptions = new HtmlPdfSaveOptions();
        pdfOptions = new HtmlPdfSaveOptions(options);
        byte[] pdf = OfficeIMO.Html.HtmlConversionDocument.Parse(scenario.Html).ToPdf(pdfOptions);
        PdfCore.PdfDocumentInfo pdfInfo = PdfCore.PdfInspector.Inspect(pdf);
        string pdfText = PdfCore.PdfReadDocument.Open(pdf).ExtractText();

        Assert.Equal(scenario.ExpectedPageCount, pdfInfo.PageCount);
        string normalizedPdfText = NormalizeCorpusWhitespace(pdfText);
        foreach (string marker in scenario.TextMarkers) {
            foreach (string searchableToken in NormalizeCorpusWhitespace(marker).Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries)) {
                Assert.Contains(searchableToken, normalizedPdfText, StringComparison.Ordinal);
            }
        }
        if (scenario.LinkUri != null) Assert.Contains(scenario.LinkUri, pdfInfo.LinkUris);
    }

    private static string NormalizeCorpusWhitespace(string value) {
        var result = new System.Text.StringBuilder(value.Length);
        bool pendingSpace = false;
        foreach (char character in value) {
            if (char.IsWhiteSpace(character)) {
                pendingSpace = result.Length > 0;
                continue;
            }
            if (pendingSpace) result.Append(' ');
            result.Append(character);
            pendingSpace = false;
        }
        return result.ToString();
    }

    private static IEnumerable<HtmlRenderVisual> EnumerateCorpusVisuals(IEnumerable<HtmlRenderVisual> visuals) {
        foreach (HtmlRenderVisual visual in visuals) {
            yield return visual;
            IEnumerable<HtmlRenderVisual>? children = visual switch {
                HtmlRenderClipGroup clip => clip.Visuals,
                HtmlRenderPathClipGroup pathClip => pathClip.Visuals,
                HtmlRenderEffectGroup effect => effect.Visuals,
                HtmlRenderSemanticGroup semantic => semantic.Visuals,
                HtmlRenderLogicalTextGroup logical => logical.Visuals,
                _ => null
            };
            if (children == null) continue;
            foreach (HtmlRenderVisual child in EnumerateCorpusVisuals(children)) yield return child;
        }
    }
}
