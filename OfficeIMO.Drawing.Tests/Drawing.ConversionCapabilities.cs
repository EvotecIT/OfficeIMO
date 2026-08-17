using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Drawing.Tests;

public sealed class DrawingConversionCapabilities {
    [Fact]
    public void SharedCatalog_HasStableUniqueRoutesAndNormalizedExtensions() {
        Assert.Equal(OfficeConversionCapabilityCatalog.All.Count,
            OfficeConversionCapabilityCatalog.All.Select(static route => route.Id).Distinct(StringComparer.Ordinal).Count());
        Assert.Equal(
            [
                "docx-pdf",
                "xlsx-pdf",
                "pptx-pdf",
                "html-pdf",
                "markdown-html",
                "html-markdown",
                "markdown-docx",
                "pdf-docx",
                "pdf-xlsx",
                "pdf-pptx",
                "pdf-html",
                "pdf-png"
            ],
            OfficeConversionCapabilityCatalog.BrowserRoutes.Select(static route => route.Id));
        Assert.All(OfficeConversionCapabilityCatalog.All, static route => {
            Assert.NotEmpty(route.SourceExtensions);
            Assert.All(route.SourceExtensions, static extension => Assert.StartsWith(".", extension, StringComparison.Ordinal));
            Assert.StartsWith(".", route.TargetExtension, StringComparison.Ordinal);
            Assert.False(string.IsNullOrWhiteSpace(route.PackageId));
            Assert.False(string.IsNullOrWhiteSpace(route.ResultContract));
        });
    }

    [Theory]
    [InlineData("onenote-pdf")]
    [InlineData("odt-pdf")]
    [InlineData("asciidoc-markdown")]
    [InlineData("markdown-latex")]
    [InlineData("visio-pdf")]
    [InlineData("pdf-docx")]
    [InlineData("mhtml-pdf")]
    public void SharedCatalog_ProtectsRepresentativeFormatFamiliesAndReverseRoutes(string routeId) {
        Assert.Contains(OfficeConversionCapabilityCatalog.All, route => route.Id == routeId);
    }

    [Fact]
    public void SharedCatalog_FiltersRoutesBySourceExtension() {
        IReadOnlyList<OfficeConversionCapability> routes =
            OfficeConversionCapabilityCatalog.FindBySourceExtension("DOCX");

        Assert.Contains(routes, static route => route.Id == "docx-pdf" && route.BrowserAvailable);
        Assert.Contains(routes, static route => route.Id == "docx-html");
        Assert.Contains(routes, static route => route.Id == "docx-markdown");
        Assert.DoesNotContain(routes, static route => route.Id == "xlsx-pdf");
    }

    [Fact]
    public void SharedCatalog_MarkdownIsDeterministicAndNamesPublicResultTypes() {
        string first = OfficeConversionCapabilityCatalog.ToMarkdown();
        string second = OfficeConversionCapabilityCatalog.ToMarkdown();

        Assert.Equal(first, second);
        Assert.Contains("| docx-pdf | DOCX | PDF | OfficeIMO.Word.Pdf |", first, StringComparison.Ordinal);
        Assert.Contains("PdfDocumentConversionResult", first, StringComparison.Ordinal);
        Assert.Contains("What it does", first, StringComparison.Ordinal);
        Assert.DoesNotContain("RtfDocument.Parse", first, StringComparison.Ordinal);
        Assert.Contains("RtfDocument.Load(stream, readOptions).ToWordDocumentResult(sourcePath)", first, StringComparison.Ordinal);
    }
}
