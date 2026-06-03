using System;
using System.IO;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfReadStreamTests {
    [Fact]
    public void RewriteApis_UseTrailerRootCatalogWhenStaleCatalogRevisionsExist() {
        byte[] pdf = BuildStaleCatalogRevisionPdf();

        PdfDocumentInfo inputInfo = PdfInspector.Inspect(pdf);
        Assert.Equal("SinglePage", inputInfo.CatalogPageLayout);
        Assert.False(inputInfo.HasReadablePageLabels);
        Assert.False(PdfInspector.Preflight(pdf).HasRewriteBlocker(PdfRewriteBlockerKind.PageLabels));

        byte[] output = PdfPageExtractor.ExtractPages(pdf, 1);

        string text = System.Text.Encoding.ASCII.GetString(output);
        Assert.Contains("/PageLayout /SinglePage", text, StringComparison.Ordinal);
        Assert.DoesNotContain("/PageLayout /TwoColumnLeft", text, StringComparison.Ordinal);
        Assert.False(PdfInspector.Inspect(output).HasReadablePageLabels);
    }

    [Fact]
    public void ReadApis_UseTrailerRootCatalogForPagesAndOutlinesWhenStaleCatalogsExist() {
        PdfDocumentInfo info = PdfInspector.Inspect(BuildStaleCatalogWithDifferentPagesAndOutlinesPdf());

        PdfPageInfo page = Assert.Single(info.Pages);
        Assert.Equal(200d, page.Width);
        Assert.Equal(200d, page.Height);
        PdfOutlineItem outline = Assert.Single(info.Outlines);
        Assert.Equal("Current", outline.Title);
        Assert.Equal(1, outline.PageNumber);
        Assert.Equal("SinglePage", info.CatalogPageLayout);
    }

    [Fact]
    public void ReadApis_UseXrefStreamRootCatalogWhenClassicTrailerIsAbsent() {
        PdfDocumentInfo info = PdfInspector.Inspect(BuildXrefStreamRootCatalogPdf());

        PdfPageInfo page = Assert.Single(info.Pages);
        Assert.Equal(200d, page.Width);
        Assert.Equal(200d, page.Height);
        PdfOutlineItem outline = Assert.Single(info.Outlines);
        Assert.Equal("Current", outline.Title);
        Assert.Equal(1, outline.PageNumber);
        Assert.Equal("SinglePage", info.CatalogPageLayout);
    }

    [Fact]
    public void RewriteApis_UseXrefStreamRootCatalogWhenClassicTrailerIsAbsent() {
        byte[] output = PdfPageExtractor.ExtractPages(BuildXrefStreamRootCatalogPdf(), 1);

        string text = System.Text.Encoding.ASCII.GetString(output);
        Assert.Contains("/PageLayout /SinglePage", text, StringComparison.Ordinal);
        Assert.DoesNotContain("/PageLayout /TwoColumnLeft", text, StringComparison.Ordinal);
        PdfOutlineItem outline = Assert.Single(PdfInspector.Inspect(output).Outlines);
        Assert.Equal("Current", outline.Title);
    }


}
