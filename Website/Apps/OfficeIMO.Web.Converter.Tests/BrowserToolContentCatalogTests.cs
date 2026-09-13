using OfficeIMO.Web.Converter.Services;
using Xunit;

namespace OfficeIMO.Web.Converter.Tests;

public sealed class BrowserToolContentCatalogTests {
    [Fact]
    public void EveryBrowserRouteHasPlainLanguageToolContent() {
        foreach (var route in ConversionRouteCatalog.All) {
            var content = BrowserToolContentCatalog.For(route);
            Assert.Equal(route.Title, content.Title);
            Assert.False(string.IsNullOrWhiteSpace(content.Summary));
            Assert.False(string.IsNullOrWhiteSpace(content.Input));
            Assert.False(string.IsNullOrWhiteSpace(content.Output));
            Assert.False(string.IsNullOrWhiteSpace(content.Expectation));
            Assert.StartsWith("/", content.GuideUrl, StringComparison.Ordinal);
            Assert.Equal(3, content.Steps.Count);
        }
    }

    [Fact]
    public void EveryPdfToolHasPlainLanguageToolContent() {
        foreach (var tool in PdfToolCatalog.All) {
            var content = BrowserToolContentCatalog.For(tool);
            Assert.Equal(tool.Label, content.Title);
            Assert.False(string.IsNullOrWhiteSpace(content.Summary));
            Assert.False(string.IsNullOrWhiteSpace(content.Input));
            Assert.False(string.IsNullOrWhiteSpace(content.Output));
            Assert.False(string.IsNullOrWhiteSpace(content.Expectation));
            Assert.StartsWith("/pdf/", content.GuideUrl);
            Assert.Equal(3, content.Steps.Count);
        }
    }

    [Fact]
    public void ProvenanceContentExplainsTheTermAndBoundary() {
        var content = BrowserToolContentCatalog.Provenance;
        Assert.Contains("Content Credentials", content.Summary, StringComparison.Ordinal);
        Assert.Contains("visible watermarks", content.Expectation, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("cannot prove", content.Expectation, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("carrier categories", content.Summary, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("Every eligible record", content.Steps[2].Description, StringComparison.OrdinalIgnoreCase);
        Assert.Equal("/provenance/", content.GuideUrl);
    }

    [Fact]
    public void CatalogExplainsEnforcedInputAndProfileLimits() {
        var textRoute = BrowserToolContentCatalog.For(ConversionRouteCatalog.Find("html-pdf"));
        var pdfRoute = BrowserToolContentCatalog.For(ConversionRouteCatalog.Find("pdf-docx"));
        var reorder = BrowserToolContentCatalog.For(PdfToolCatalog.Find("reorder"));
        var optimize = BrowserToolContentCatalog.For(PdfToolCatalog.Find("optimize"));

        Assert.Contains("500,000 characters", textRoute.Input, StringComparison.Ordinal);
        Assert.Contains("500 pages", pdfRoute.Input, StringComparison.Ordinal);
        Assert.Contains("rejects repeated or omitted pages", reorder.Expectation, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("Balanced and archival", optimize.Output, StringComparison.Ordinal);
        Assert.Contains("larger file", optimize.Expectation, StringComparison.OrdinalIgnoreCase);
    }
}
