using System;
using System.IO;
using System.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfInspectorTests {
    [Fact]
    public void Preflight_AllowsCatalogViewSettingPdfReadAndRewrite() {
        PdfDocumentPreflight report = PdfInspector.Preflight(BuildCatalogViewSettingPdf());

        Assert.True(report.CanRead);
        Assert.True(report.CanRewrite);
        Assert.True(report.Probe.HasCatalogViewSettings);
        Assert.False(report.Probe.HasOutlines);
        Assert.NotNull(report.DocumentInfo);
        Assert.True(report.DocumentInfo!.HasCatalogViewSettings);
        Assert.False(report.DocumentInfo.HasOutlines);
        Assert.Equal("FullScreen", report.DocumentInfo.CatalogPageMode);
        Assert.Equal("TwoColumnLeft", report.DocumentInfo.CatalogPageLayout);
        Assert.Null(report.DocumentInfo.CatalogVersion);
        Assert.Null(report.DocumentInfo.CatalogLanguage);
        Assert.Empty(report.DocumentInfo.Outlines);
        Assert.Empty(report.ReadBlockers);
        Assert.Empty(report.Diagnostics);
        Assert.Empty(report.RewriteBlockers);
        Assert.False(report.HasRewriteBlocker(PdfRewriteBlockerKind.CatalogViewSettings));
    }

    [Fact]
    public void Preflight_ReadsCatalogLanguageAndVersion() {
        PdfDocumentPreflight report = PdfInspector.Preflight(BuildCatalogIdentityPdf());

        Assert.True(report.CanRead);
        Assert.True(report.CanRewrite);
        Assert.NotNull(report.DocumentInfo);
        Assert.Equal("1.7", report.DocumentInfo!.CatalogVersion);
        Assert.Equal("pl-PL", report.DocumentInfo.CatalogLanguage);
        Assert.Null(report.DocumentInfo.CatalogPageMode);
        Assert.Null(report.DocumentInfo.CatalogPageLayout);
        Assert.Empty(report.ReadBlockers);
        Assert.Empty(report.Diagnostics);
        Assert.Empty(report.RewriteBlockers);
    }

    [Fact]
    public void Preflight_AllowsPageLabelPdfReadAndRewrite() {
        PdfDocumentPreflight report = PdfInspector.Preflight(BuildPageLabelPdf());

        Assert.True(report.CanRead);
        Assert.True(report.CanRewrite);
        Assert.True(report.Probe.HasPageLabels);
        Assert.False(report.Probe.HasCatalogNameTrees);
        Assert.False(report.Probe.HasNamedDestinations);
        Assert.NotNull(report.DocumentInfo);
        Assert.True(report.DocumentInfo!.HasPageLabels);
        Assert.False(report.DocumentInfo.HasCatalogNameTrees);
        Assert.False(report.DocumentInfo.HasNamedDestinations);
        AssertPageLabel(report.DocumentInfo, 0, 1, "D", null, 1);
        Assert.Empty(report.ReadBlockers);
        Assert.Empty(report.Diagnostics);
        Assert.Empty(report.RewriteBlockers);
        Assert.False(report.HasRewriteBlocker(PdfRewriteBlockerKind.PageLabels));
    }

    [Fact]
    public void Preflight_AllowsComplexPageLabelPdfReadButBlocksRewrite() {
        PdfDocumentPreflight report = PdfInspector.Preflight(BuildComplexPageLabelPdf());

        Assert.True(report.CanRead);
        Assert.False(report.CanRewrite);
        Assert.True(report.Probe.HasPageLabels);
        Assert.NotNull(report.DocumentInfo);
        Assert.True(report.DocumentInfo!.HasPageLabels);
        Assert.False(report.DocumentInfo.HasReadablePageLabels);
        Assert.Equal(0, report.DocumentInfo.PageLabelCount);
        Assert.Empty(report.DocumentInfo.PageLabels);
        Assert.Empty(report.ReadBlockers);
        Assert.Contains(report.Diagnostics, diagnostic => diagnostic.Contains("page labels", StringComparison.OrdinalIgnoreCase));
        Assert.Contains(report.RewriteBlockers, blocker => blocker.Kind == PdfRewriteBlockerKind.PageLabels);
        Assert.True(report.HasRewriteBlocker(PdfRewriteBlockerKind.PageLabels));
    }


}
