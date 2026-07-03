using System;
using System.Text;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfDocumentVisualQualityTests {
    [Fact]
    public void CatalogUriBase_CanBeEmittedClonedProbedAndPreservedOnExtraction() {
        var options = new PdfOptions().SetCatalogUriBase("https://evotec.xyz/docs/");

        byte[] bytes = PdfDocument.Create(options)
            .CatalogUriBase("https://officeimo.net/pdf/")
            .Paragraph(p => p.Text("Catalog URI base proof."))
            .PageBreak()
            .Paragraph(p => p.Text("Second page."))
            .ToBytes();

        string raw = Encoding.ASCII.GetString(bytes);
        PdfDocumentInfo info = PdfInspector.Inspect(bytes);
        PdfDocumentPreflight preflight = PdfInspector.Preflight(bytes);
        PdfOptions clone = options.Clone();

        Assert.Contains("/URI << /Base (https://officeimo.net/pdf/) >>", raw, StringComparison.Ordinal);
        Assert.True(info.HasCatalogUri);
        Assert.True(preflight.Probe.HasCatalogUri);
        Assert.True(preflight.CanRewrite);
        Assert.Equal("https://evotec.xyz/docs/", clone.CatalogUriBase);
        Assert.NotNull(typeof(PdfOptions).GetMethod("SetCatalogUriBase", new[] { typeof(string) }));
        Assert.NotNull(typeof(PdfDocument).GetMethod("CatalogUriBase", new[] { typeof(string) }));

        byte[] extracted = PdfPageExtractor.ExtractPages(bytes, 1);
        string extractedRaw = Encoding.ASCII.GetString(extracted);
        Assert.Contains("/URI << /Base (https://officeimo.net/pdf/) >>", extractedRaw, StringComparison.Ordinal);
        Assert.True(PdfInspector.Probe(extracted).HasCatalogUri);

        Assert.Null(new PdfOptions().SetCatalogUriBase("https://evotec.xyz/").ClearCatalogUriBase().CatalogUriBase);
        Assert.Throws<ArgumentException>(() => new PdfOptions().SetCatalogUriBase(""));
        Assert.Throws<ArgumentException>(() => new PdfOptions().SetCatalogUriBase("relative/path"));
        Assert.Throws<ArgumentException>(() => PdfDocument.Create().CatalogUriBase("bad\u0001uri"));
    }

    [Fact]
    public void CatalogUriBase_AllowsRelativeUriActionsAndPreservesThemOnExtraction() {
        const string paragraphTarget = "assets/report.html#summary";
        const string headingTarget = "appendix/index.html";
        const string tableTarget = "tables/detail.html";

        byte[] bytes = PdfDocument.Create(new PdfOptions().SetCatalogUriBase("https://officeimo.net/docs/"))
            .Paragraph(p => p.Link("relative paragraph", paragraphTarget, contents: "Relative paragraph"))
            .H2("relative heading", linkUri: headingTarget, linkContents: "Relative heading")
            .TableWithLinks(
                new[] {
                    new[] { "Cell", "Value" }
                },
                new() {
                    [(0, 1)] = tableTarget
                })
            .ToBytes();

        string raw = Encoding.ASCII.GetString(bytes);
        PdfDocumentInfo info = PdfInspector.Inspect(bytes);
        PdfLogicalDocument logical = PdfLogicalDocument.Load(bytes);

        Assert.Contains("/URI << /Base (https://officeimo.net/docs/) >>", raw, StringComparison.Ordinal);
        Assert.Contains("/URI (assets/report.html#summary)", raw, StringComparison.Ordinal);
        Assert.Contains("/URI (appendix/index.html)", raw, StringComparison.Ordinal);
        Assert.Contains("/URI (tables/detail.html)", raw, StringComparison.Ordinal);
        Assert.True(info.HasCatalogUri);
        Assert.Contains(paragraphTarget, info.LinkUris);
        Assert.Contains(headingTarget, info.LinkUris);
        Assert.Contains(tableTarget, info.LinkUris);
        Assert.Single(logical.GetLinksByUri(paragraphTarget));

        byte[] extracted = PdfPageExtractor.ExtractPages(bytes, 1);
        PdfDocumentInfo extractedInfo = PdfInspector.Inspect(extracted);
        Assert.True(extractedInfo.HasCatalogUri);
        Assert.Contains(paragraphTarget, extractedInfo.LinkUris);
        Assert.Contains(headingTarget, extractedInfo.LinkUris);
        Assert.Contains(tableTarget, extractedInfo.LinkUris);
    }

    [Fact]
    public void RelativeUriActionsRequireCatalogUriBaseWhenRendering() {
        var exception = Assert.Throws<ArgumentException>(() =>
            PdfDocument.Create()
                .Paragraph(p => p.Link("relative paragraph", "assets/report.html"))
                .ToBytes());

        Assert.Contains("Relative PDF URI link targets require PdfOptions.CatalogUriBase.", exception.Message, StringComparison.Ordinal);
    }
}
