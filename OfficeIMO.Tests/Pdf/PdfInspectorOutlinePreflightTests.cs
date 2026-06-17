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

    [Fact]
    public void ReadApis_LimitWideOutlineTraversal() {
        PdfDocumentInfo info = PdfInspector.Inspect(BuildWideOutlinePdf(2_100));

        Assert.Equal(2_048, info.Outlines.Count);
        Assert.Equal("Item 0", info.Outlines[0].Title);
        Assert.Equal("Item 2047", info.Outlines[2_047].Title);
    }

    [Fact]
    public void ReadApis_LimitDeepOutlineTraversal() {
        PdfDocumentInfo info = PdfInspector.Inspect(BuildDeepOutlinePdf(80));

        PdfOutlineItem current = Assert.Single(info.Outlines);
        int depth = 1;
        while (current.Children.Count > 0) {
            current = Assert.Single(current.Children);
            depth++;
        }

        Assert.Equal(64, depth);
        Assert.Equal("Depth 64", current.Title);
    }


}
