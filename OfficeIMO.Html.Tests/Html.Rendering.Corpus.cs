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

    [Fact]
    public void HtmlRenderingCorpus_StaticStandardsGridUsesTwoAuthoredColumns() {
        HtmlRenderingCorpusCase scenario = HtmlRenderingCorpus.All.Single(item => item.Id == "static-standards-showcase");
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(scenario.Html, scenario.CreateOptions());
        HtmlRenderText firstRow = Assert.Single(
            rendered.Pages.SelectMany(page => EnumerateCorpusVisuals(page.Scene)).OfType<HtmlRenderText>(),
            text => text.Text == "Inherited row A");
        HtmlRenderText badge = Assert.Single(
            rendered.Pages.SelectMany(page => EnumerateCorpusVisuals(page.Scene)).OfType<HtmlRenderText>(),
            text => text.Text == "Clipped vector badge");
        HtmlRenderText evidence = Assert.Single(
            rendered.Pages.SelectMany(page => EnumerateCorpusVisuals(page.Scene)).OfType<HtmlRenderText>(),
            text => text.Text == "Named page evidence");

        Assert.True(firstRow.X < badge.X);
        Assert.InRange(Math.Abs(badge.X - evidence.X), 0D, 2D);
    }

    [Fact]
    public void HtmlRenderingCorpus_StaticStandardsRunningHeaderPaintsOnEveryRasterPage() {
        HtmlRenderingCorpusCase scenario = HtmlRenderingCorpus.All.Single(item => item.Id == "static-standards-showcase");
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(scenario.Html, scenario.CreateOptions());

        Assert.Equal(2, rendered.Pages.Count);
        foreach (HtmlRenderPage page in rendered.Pages) {
            OfficeDrawing drawing = page.CreateDrawing();
            OfficeDrawingText header = Assert.Single(
                drawing.Elements.OfType<OfficeDrawingText>(),
                text => text.Text.Contains("Managed static standards", StringComparison.Ordinal));
            OfficeRasterImage image = OfficeDrawingRasterRenderer.Render(drawing, 1D, OfficeColor.White);
            int coloredPixels = 0;
            int left = Math.Max(0, (int)Math.Floor(header.X - 5D));
            int top = Math.Max(0, (int)Math.Floor(header.Y - 5D));
            int right = Math.Min(image.Width - 1, (int)Math.Ceiling(header.X + header.Width + 5D));
            int bottom = Math.Min(image.Height - 1, (int)Math.Ceiling(header.Y + header.Height + 5D));
            for (int y = top; y <= bottom; y++) {
                for (int x = left; x <= right; x++) {
                    OfficeColor pixel = image.GetPixel(x, y);
                    if (pixel.B > pixel.R + 20 && pixel.B > pixel.G + 5) coloredPixels++;
                }
            }

            Assert.True(coloredPixels > 20, $"Page {page.PageNumber} running header produced only {coloredPixels} blue raster pixels.");
        }
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
        string svgText = System.Net.WebUtility.HtmlDecode(
            System.Text.RegularExpressions.Regex.Replace(svg, "<[^>]+>", string.Empty));
        foreach (string word in NormalizeCorpusWhitespace(scenario.TextMarkers[0]).Split(' ')) {
            Assert.True(
                svg.Contains(word, StringComparison.Ordinal) || svgText.Contains(word, StringComparison.Ordinal),
                $"Expected SVG paint text to retain '{word}' either in one text node or across positioned grapheme nodes.");
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

        WriteStaticStandardsReviewArtifacts(scenario, rendered, pdf);
    }

    private static void WriteStaticStandardsReviewArtifacts(
        HtmlRenderingCorpusCase scenario,
        HtmlRenderDocument rendered,
        byte[] pdf) {
        string? directory = Environment.GetEnvironmentVariable("OFFICEIMO_HTML_STANDARDS_ARTIFACT_DIR");
        if (scenario.Id != "static-standards-showcase" || string.IsNullOrWhiteSpace(directory)) return;

        Directory.CreateDirectory(directory);
        File.WriteAllBytes(Path.Combine(directory, "static-standards.pdf"), pdf);
        for (int pageIndex = 0; pageIndex < rendered.Pages.Count; pageIndex++) {
            OfficeDrawing pageDrawing = rendered.Pages[pageIndex].CreateDrawing();
            byte[] pagePng = OfficeDrawingRasterRenderer.ToPng(
                pageDrawing,
                1D,
                OfficeColor.White);
            File.WriteAllBytes(
                Path.Combine(directory, "static-standards-page-" + (pageIndex + 1).ToString(System.Globalization.CultureInfo.InvariantCulture) + ".png"),
                pagePng);
            File.WriteAllText(
                Path.Combine(directory, "static-standards-page-" + (pageIndex + 1).ToString(System.Globalization.CultureInfo.InvariantCulture) + ".svg"),
                OfficeDrawingSvgExporter.ToSvg(pageDrawing, 1D));
        }
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
