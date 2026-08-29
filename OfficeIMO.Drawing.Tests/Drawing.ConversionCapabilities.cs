using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Drawing.Tests;

public sealed class DrawingConversionCapabilities {
    [Fact]
    public void CapabilityCatalog_PreservesTheOriginalPublicConstructorSignature() {
        Type[] originalParameters = [
            typeof(string),
            typeof(string),
            typeof(string),
            typeof(OfficeConversionInputKind),
            typeof(IEnumerable<string>),
            typeof(string),
            typeof(string),
            typeof(string),
            typeof(string),
            typeof(OfficeConversionFidelityKind),
            typeof(string),
            typeof(bool),
            typeof(bool)
        ];

        Assert.NotNull(typeof(OfficeConversionCapability).GetConstructor(originalParameters));

        Type[] originalExplicitAssessmentParameters = [
            typeof(string),
            typeof(string),
            typeof(string),
            typeof(OfficeConversionInputKind),
            typeof(IEnumerable<string>),
            typeof(string),
            typeof(string),
            typeof(string),
            typeof(string),
            typeof(OfficeConversionFidelityKind),
            typeof(string),
            typeof(bool),
            typeof(bool),
            typeof(OfficeConversionSupportLevel),
            typeof(string),
            typeof(string)
        ];

        Assert.NotNull(typeof(OfficeConversionCapability).GetConstructor(originalExplicitAssessmentParameters));
    }

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
            Assert.False(string.IsNullOrWhiteSpace(route.SupportEvidence));
            Assert.False(string.IsNullOrWhiteSpace(route.KnownLimitations));
        });
        Assert.DoesNotContain(
            OfficeConversionCapabilityCatalog.All,
            static route => route.SupportLevel == OfficeConversionSupportLevel.ReferenceVerified);
        Assert.Equal(8, OfficeConversionCapabilityCatalog.All.Count(static route => route.SupportLevel == OfficeConversionSupportLevel.Advanced));
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
        Assert.Contains("| Support | Evidence | Known limits |", first, StringComparison.Ordinal);
        Assert.Contains("| docx-pdf | DOCX | PDF | OfficeIMO.Word.Pdf | FixedLayout | FixedLayoutAppearance |", first, StringComparison.Ordinal);
        Assert.Contains("| Advanced | Realistic DOCX fixtures", first, StringComparison.Ordinal);
        Assert.DoesNotContain("RtfDocument.Parse", first, StringComparison.Ordinal);
        Assert.Contains("RtfDocument.Load(stream, readOptions).ToWordDocumentResult(sourcePath)", first, StringComparison.Ordinal);
    }

    [Fact]
    public void SharedCatalog_SeparatesOutputModelFromSupportDepth() {
        OfficeConversionCapability pdfToWord = Assert.IsType<OfficeConversionCapability>(
            OfficeConversionCapabilityCatalog.Find("pdf-docx"));
        OfficeConversionCapability pdfToPng = Assert.IsType<OfficeConversionCapability>(
            OfficeConversionCapabilityCatalog.Find("pdf-png"));

        Assert.Equal(OfficeConversionFidelityKind.Editable, pdfToWord.Fidelity);
        Assert.Equal(OfficeConversionSupportLevel.Targeted, pdfToWord.SupportLevel);
        Assert.Equal(OfficeConversionFidelityKind.FixedLayout, pdfToPng.Fidelity);
        Assert.Equal(OfficeConversionSupportLevel.Advanced, pdfToPng.SupportLevel);
        Assert.Contains("not page-layout recovery", pdfToWord.KnownLimitations, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void SharedCatalog_OneNoteImageRoutesExposeACompleteFilePipeline() {
        OfficeConversionCapability route = Assert.IsType<OfficeConversionCapability>(
            OfficeConversionCapabilityCatalog.Find("onenote-png"));

        Assert.Equal(OfficeConversionInputKind.File, route.InputKind);
        Assert.Contains("OneNoteSectionReader.Read(stream)", route.Api, StringComparison.Ordinal);
        Assert.Contains(".ExportImages(", route.Api, StringComparison.Ordinal);
    }
}
