using System;
using System.IO;
using System.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfInspectorTests {
    [Fact]
    public void Preflight_AllowsDirectNamedDestinationPdfReadAndRewrite() {
        PdfDocumentPreflight report = PdfInspector.Preflight(BuildNamedDestinationPdf());

        Assert.True(report.CanRead);
        Assert.True(report.CanRewrite);
        Assert.True(report.Probe.HasNamedDestinations);
        Assert.False(report.Probe.HasCatalogNameTrees);
        Assert.False(report.Probe.HasPageLabels);
        Assert.NotNull(report.DocumentInfo);
        Assert.True(report.DocumentInfo!.HasNamedDestinations);
        Assert.False(report.DocumentInfo.HasCatalogNameTrees);
        Assert.False(report.DocumentInfo.HasPageLabels);
        AssertNamedDestination(report.DocumentInfo, "Chapter1", 1, 200);
        Assert.Empty(report.ReadBlockers);
        Assert.Empty(report.Diagnostics);
        Assert.Empty(report.RewriteBlockers);
        Assert.False(report.HasRewriteBlocker(PdfRewriteBlockerKind.NamedDestinations));
    }

    [Fact]
    public void Preflight_AllowsNamedDestinationNameTreeReadAndRewrite() {
        PdfDocumentPreflight report = PdfInspector.Preflight(BuildNamedDestinationNameTreePdf());

        Assert.True(report.CanRead);
        Assert.True(report.CanRewrite);
        Assert.True(report.Probe.HasNamedDestinations);
        Assert.True(report.Probe.HasCatalogNameTrees);
        Assert.NotNull(report.DocumentInfo);
        Assert.True(report.DocumentInfo!.HasNamedDestinations);
        Assert.True(report.DocumentInfo.HasCatalogNameTrees);
        AssertNamedDestination(report.DocumentInfo, "Chapter1", 1, 200);
        Assert.Empty(report.ReadBlockers);
        Assert.Empty(report.Diagnostics);
        Assert.Empty(report.RewriteBlockers);
        Assert.False(report.HasRewriteBlocker(PdfRewriteBlockerKind.NamedDestinations));
    }

    [Fact]
    public void Preflight_AllowsNamedDestinationNameTreeKidsReadAndRewrite() {
        PdfDocumentPreflight report = PdfInspector.Preflight(BuildNamedDestinationNameTreeWithKidsPdf());

        Assert.True(report.CanRead);
        Assert.True(report.CanRewrite);
        Assert.True(report.Probe.HasNamedDestinations);
        Assert.True(report.Probe.HasCatalogNameTrees);
        Assert.NotNull(report.DocumentInfo);
        Assert.True(report.DocumentInfo!.HasNamedDestinations);
        Assert.True(report.DocumentInfo.HasCatalogNameTrees);
        AssertNamedDestination(report.DocumentInfo, "Chapter1", 1, 200);
        Assert.Empty(report.ReadBlockers);
        Assert.Empty(report.Diagnostics);
        Assert.Empty(report.RewriteBlockers);
        Assert.False(report.HasRewriteBlocker(PdfRewriteBlockerKind.NamedDestinations));
    }

    [Fact]
    public void Preflight_HonorsConfiguredNamedDestinationRewriteTraversalLimits() {
        byte[] pdf = BuildDeepNamedDestinationNameTreePdf();
        var nodeLimitedOptions = new PdfReadOptions {
            Limits = new PdfReadLimits { MaxNameTreeNodes = 2 }
        };
        var depthLimitedOptions = new PdfReadOptions {
            Limits = new PdfReadLimits { MaxNameTreeDepth = 1 }
        };

        PdfDocumentPreflight nodeLimited = PdfInspector.Preflight(pdf, nodeLimitedOptions);
        PdfDocumentPreflight depthLimited = PdfInspector.Preflight(pdf, depthLimitedOptions);

        Assert.True(nodeLimited.HasRewriteBlocker(PdfRewriteBlockerKind.NamedDestinations));
        Assert.True(depthLimited.HasRewriteBlocker(PdfRewriteBlockerKind.NamedDestinations));
    }

    [Fact]
    public void Preflight_AllowsUnsupportedNamedDestinationNameTreeReadButBlocksRewrite() {
        PdfDocumentPreflight report = PdfInspector.Preflight(BuildComplexNamedDestinationNameTreePdf());

        Assert.True(report.CanRead);
        Assert.False(report.CanRewrite);
        Assert.True(report.Probe.HasNamedDestinations);
        Assert.True(report.Probe.HasCatalogNameTrees);
        Assert.NotNull(report.DocumentInfo);
        Assert.True(report.DocumentInfo!.HasNamedDestinations);
        Assert.True(report.DocumentInfo.HasCatalogNameTrees);
        Assert.Empty(report.ReadBlockers);
        Assert.Contains("PDF named destinations are not supported for rewriting by OfficeIMO.Pdf yet.", report.Diagnostics);
        AssertRewriteBlocker(report, PdfRewriteBlockerKind.NamedDestinations, "PDF named destinations are not supported for rewriting by OfficeIMO.Pdf yet.");
    }

    [Fact]
    public void Preflight_BlocksMalformedNamedDestinationNameTreeShapes() {
        AssertMalformedNameTree(BuildMixedNamedDestinationNameTreePdf());
        AssertMalformedNameTree(BuildDirectKidNamedDestinationNameTreePdf());

        static void AssertMalformedNameTree(byte[] pdf) {
            PdfDocumentPreflight report = PdfInspector.Preflight(pdf);

            Assert.True(report.CanRead);
            Assert.False(report.CanRewrite);
            Assert.True(report.Probe.HasNamedDestinations);
            Assert.True(report.Probe.HasCatalogNameTrees);
            Assert.Empty(report.ReadBlockers);
            Assert.Contains("PDF named destinations are not supported for rewriting by OfficeIMO.Pdf yet.", report.Diagnostics);
            AssertRewriteBlocker(report, PdfRewriteBlockerKind.NamedDestinations, "PDF named destinations are not supported for rewriting by OfficeIMO.Pdf yet.");
        }
    }

    [Fact]
    public void Preflight_AllowsDestinationOpenActionPdfReadAndRewrite() {
        PdfDocumentPreflight report = PdfInspector.Preflight(BuildOpenActionPdf());

        Assert.True(report.CanRead);
        Assert.True(report.CanRewrite);
        Assert.True(report.Probe.HasOpenActions);
        Assert.False(report.Probe.HasViewerPreferences);
        Assert.NotNull(report.DocumentInfo);
        Assert.True(report.DocumentInfo!.HasOpenActions);
        Assert.False(report.DocumentInfo.HasViewerPreferences);
        AssertOpenAction(report.DocumentInfo, "Destination", 1, null);
        Assert.Empty(report.ReadBlockers);
        Assert.Empty(report.Diagnostics);
        Assert.Empty(report.RewriteBlockers);
        Assert.False(report.HasRewriteBlocker(PdfRewriteBlockerKind.OpenActions));
    }

    [Fact]
    public void Preflight_AllowsGoToOpenActionDictionaryReadAndRewrite() {
        PdfDocumentPreflight report = PdfInspector.Preflight(BuildOpenActionDictionaryPdf());

        Assert.True(report.CanRead);
        Assert.True(report.CanRewrite);
        Assert.True(report.Probe.HasOpenActions);
        Assert.NotNull(report.DocumentInfo);
        Assert.True(report.DocumentInfo!.HasOpenActions);
        AssertOpenAction(report.DocumentInfo, "GoTo", 1, null);
        Assert.Empty(report.ReadBlockers);
        Assert.Empty(report.Diagnostics);
        Assert.Empty(report.RewriteBlockers);
        Assert.False(report.HasRewriteBlocker(PdfRewriteBlockerKind.OpenActions));
    }

    [Fact]
    public void Preflight_AllowsComplexOpenActionDictionaryReadButBlocksRewrite() {
        PdfDocumentPreflight report = PdfInspector.Preflight(BuildComplexOpenActionDictionaryPdf());

        Assert.True(report.CanRead);
        Assert.False(report.CanRewrite);
        Assert.True(report.Probe.HasOpenActions);
        Assert.NotNull(report.DocumentInfo);
        Assert.True(report.DocumentInfo!.HasOpenActions);
        Assert.False(report.DocumentInfo.HasReadableOpenAction);
        Assert.Null(report.DocumentInfo.OpenAction);
        Assert.Empty(report.ReadBlockers);
        Assert.Contains("PDF open actions are not supported for rewriting by OfficeIMO.Pdf yet.", report.Diagnostics);
        AssertRewriteBlocker(report, PdfRewriteBlockerKind.OpenActions, "PDF open actions are not supported for rewriting by OfficeIMO.Pdf yet.");
    }


}
