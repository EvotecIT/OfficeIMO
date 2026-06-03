using System;
using System.IO;
using System.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfInspectorTests {
    [Fact]
    public void Preflight_AllowsSimpleOutlinePdfReadAndRewrite() {
        PdfDocumentPreflight report = PdfInspector.Preflight(BuildOutlinePdf());

        Assert.True(report.CanRead);
        Assert.True(report.CanRewrite);
        Assert.True(report.Probe.HasOutlines);
        Assert.True(report.Probe.HasCatalogViewSettings);
        Assert.NotNull(report.DocumentInfo);
        Assert.True(report.DocumentInfo!.HasOutlines);
        Assert.True(report.DocumentInfo.HasCatalogViewSettings);
        Assert.NotEmpty(report.DocumentInfo.Outlines);
        Assert.Empty(report.ReadBlockers);
        Assert.Empty(report.Diagnostics);
        Assert.Empty(report.RewriteBlockers);
        Assert.False(report.HasRewriteBlocker(PdfRewriteBlockerKind.CatalogViewSettings));
    }

    [Fact]
    public void Preflight_AllowsComplexOutlinePdfReadButBlocksRewrite() {
        PdfDocumentPreflight report = PdfInspector.Preflight(BuildComplexOutlinePdf());

        Assert.True(report.CanRead);
        Assert.False(report.CanRewrite);
        Assert.True(report.Probe.HasOutlines);
        Assert.True(report.Probe.HasCatalogViewSettings);
        Assert.NotNull(report.DocumentInfo);
        Assert.True(report.DocumentInfo!.HasOutlines);
        Assert.True(report.DocumentInfo.HasCatalogViewSettings);
        Assert.NotEmpty(report.DocumentInfo.Outlines);
        Assert.Empty(report.ReadBlockers);
        Assert.Contains("PDF outlines are not supported for rewriting by OfficeIMO.Pdf yet.", report.Diagnostics);
        AssertRewriteBlocker(report, PdfRewriteBlockerKind.Outlines, "PDF outlines are not supported for rewriting by OfficeIMO.Pdf yet.");
        Assert.False(report.HasRewriteBlocker(PdfRewriteBlockerKind.CatalogViewSettings));
    }

    [Fact]
    public void Preflight_BlocksCyclicGoToOutlineDestinationsWithoutRecursingForever() {
        PdfDocumentPreflight report = PdfInspector.Preflight(BuildCyclicGoToOutlineDestinationPdf());

        Assert.True(report.CanRead);
        Assert.False(report.CanRewrite);
        Assert.True(report.Probe.HasOutlines);
        Assert.NotNull(report.DocumentInfo);
        PdfOutlineItem outline = Assert.Single(report.DocumentInfo!.Outlines);
        Assert.Equal("Cyclic", outline.Title);
        Assert.Null(outline.PageNumber);
        Assert.Empty(report.ReadBlockers);
        AssertRewriteBlocker(report, PdfRewriteBlockerKind.Outlines, "PDF outlines are not supported for rewriting by OfficeIMO.Pdf yet.");
    }


}
