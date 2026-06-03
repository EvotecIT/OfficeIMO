using System;
using System.IO;
using System.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfInspectorTests {
    [Fact]
    public void Preflight_AllowsSimpleViewerPreferencePdfReadAndRewrite() {
        PdfDocumentPreflight report = PdfInspector.Preflight(BuildViewerPreferencePdf());

        Assert.True(report.CanRead);
        Assert.True(report.CanRewrite);
        Assert.True(report.Probe.HasViewerPreferences);
        Assert.False(report.Probe.HasOpenActions);
        Assert.NotNull(report.DocumentInfo);
        Assert.True(report.DocumentInfo!.HasViewerPreferences);
        Assert.False(report.DocumentInfo.HasOpenActions);
        AssertViewerPreferences(report.DocumentInfo);
        Assert.Empty(report.ReadBlockers);
        Assert.Empty(report.RewriteBlockers);
        Assert.DoesNotContain("PDF viewer preferences are not supported for rewriting by OfficeIMO.Pdf yet.", report.Diagnostics);
    }

    [Fact]
    public void Preflight_AllowsComplexViewerPreferencePdfReadButBlocksRewrite() {
        PdfDocumentPreflight report = PdfInspector.Preflight(BuildComplexViewerPreferencePdf());

        Assert.True(report.CanRead);
        Assert.False(report.CanRewrite);
        Assert.True(report.Probe.HasViewerPreferences);
        Assert.NotNull(report.DocumentInfo);
        Assert.True(report.DocumentInfo!.HasViewerPreferences);
        Assert.False(report.DocumentInfo.HasReadableViewerPreferences);
        Assert.Null(report.DocumentInfo.ViewerPreferences);
        Assert.Empty(report.ReadBlockers);
        Assert.Contains("PDF viewer preferences are not supported for rewriting by OfficeIMO.Pdf yet.", report.Diagnostics);
        AssertRewriteBlocker(report, PdfRewriteBlockerKind.ViewerPreferences, "PDF viewer preferences are not supported for rewriting by OfficeIMO.Pdf yet.");
    }

    [Fact]
    public void Preflight_AllowsTaggedPdfReadButBlocksRewrite() {
        PdfDocumentPreflight report = PdfInspector.Preflight(BuildTaggedPdf());

        Assert.True(report.CanRead);
        Assert.False(report.CanRewrite);
        Assert.True(report.Probe.HasTaggedContent);
        Assert.NotNull(report.DocumentInfo);
        Assert.True(report.DocumentInfo!.HasTaggedContent);
        Assert.Empty(report.ReadBlockers);
        Assert.Contains("PDF tagged content structure is not supported for rewriting by OfficeIMO.Pdf yet.", report.Diagnostics);
        AssertRewriteBlocker(report, PdfRewriteBlockerKind.TaggedContent, "PDF tagged content structure is not supported for rewriting by OfficeIMO.Pdf yet.");
    }

    [Fact]
    public void Preflight_AllowsSimpleXmpMetadataPdfReadAndRewrite() {
        PdfDocumentPreflight report = PdfInspector.Preflight(BuildXmpMetadataPdf());

        Assert.True(report.CanRead);
        Assert.True(report.CanRewrite);
        Assert.True(report.Probe.HasXmpMetadata);
        Assert.False(report.Probe.HasCatalogUri);
        Assert.False(report.Probe.HasOutputIntents);
        Assert.NotNull(report.DocumentInfo);
        Assert.True(report.DocumentInfo!.HasXmpMetadata);
        Assert.False(report.DocumentInfo.HasCatalogUri);
        Assert.False(report.DocumentInfo.HasOutputIntents);
        Assert.Empty(report.ReadBlockers);
        Assert.Empty(report.RewriteBlockers);
        Assert.DoesNotContain("PDF XMP metadata is not supported for rewriting by OfficeIMO.Pdf yet.", report.Diagnostics);
    }

    [Fact]
    public void Preflight_AllowsComplexXmpMetadataPdfReadButBlocksRewrite() {
        PdfDocumentPreflight report = PdfInspector.Preflight(BuildComplexXmpMetadataPdf());

        Assert.True(report.CanRead);
        Assert.False(report.CanRewrite);
        Assert.True(report.Probe.HasXmpMetadata);
        Assert.NotNull(report.DocumentInfo);
        Assert.True(report.DocumentInfo!.HasXmpMetadata);
        Assert.Empty(report.ReadBlockers);
        Assert.Contains("PDF XMP metadata is not supported for rewriting by OfficeIMO.Pdf yet.", report.Diagnostics);
        AssertRewriteBlocker(report, PdfRewriteBlockerKind.XmpMetadata, "PDF XMP metadata is not supported for rewriting by OfficeIMO.Pdf yet.");
    }

    [Fact]
    public void Preflight_AllowsSimpleCatalogUriPdfReadAndRewrite() {
        PdfDocumentPreflight report = PdfInspector.Preflight(BuildCatalogUriPdf());

        Assert.True(report.CanRead);
        Assert.True(report.CanRewrite);
        Assert.True(report.Probe.HasCatalogUri);
        Assert.False(report.Probe.HasOutputIntents);
        Assert.NotNull(report.DocumentInfo);
        Assert.True(report.DocumentInfo!.HasCatalogUri);
        Assert.False(report.DocumentInfo.HasOutputIntents);
        Assert.Empty(report.ReadBlockers);
        Assert.Empty(report.RewriteBlockers);
        Assert.DoesNotContain("PDF catalog URI dictionaries are not supported for rewriting by OfficeIMO.Pdf yet.", report.Diagnostics);
    }

    [Fact]
    public void Preflight_DoesNotTreatLinkAnnotationUriAsCatalogUri() {
        PdfDocumentPreflight report = PdfInspector.Preflight(BuildAnnotatedPdf());

        Assert.True(report.CanRead);
        Assert.True(report.CanRewrite);
        Assert.True(report.Probe.HasAnnotations);
        Assert.False(report.Probe.HasCatalogUri);
        Assert.NotNull(report.DocumentInfo);
        Assert.True(report.DocumentInfo!.HasAnnotations);
        Assert.False(report.DocumentInfo.HasCatalogUri);
        Assert.Empty(report.ReadBlockers);
        Assert.Empty(report.RewriteBlockers);
    }

    [Fact]
    public void Preflight_AllowsComplexCatalogUriPdfReadButBlocksRewrite() {
        PdfDocumentPreflight report = PdfInspector.Preflight(BuildComplexCatalogUriPdf());

        Assert.True(report.CanRead);
        Assert.False(report.CanRewrite);
        Assert.True(report.Probe.HasCatalogUri);
        Assert.NotNull(report.DocumentInfo);
        Assert.True(report.DocumentInfo!.HasCatalogUri);
        Assert.Empty(report.ReadBlockers);
        Assert.Contains("PDF catalog URI dictionaries are not supported for rewriting by OfficeIMO.Pdf yet.", report.Diagnostics);
        AssertRewriteBlocker(report, PdfRewriteBlockerKind.CatalogUri, "PDF catalog URI dictionaries are not supported for rewriting by OfficeIMO.Pdf yet.");
    }

    [Fact]
    public void Preflight_AllowsSimpleOutputIntentPdfReadAndRewrite() {
        PdfDocumentPreflight report = PdfInspector.Preflight(BuildOutputIntentPdf());

        Assert.True(report.CanRead);
        Assert.True(report.CanRewrite);
        Assert.True(report.Probe.HasOutputIntents);
        Assert.False(report.Probe.HasXmpMetadata);
        Assert.False(report.Probe.HasCatalogUri);
        Assert.NotNull(report.DocumentInfo);
        Assert.True(report.DocumentInfo!.HasOutputIntents);
        Assert.False(report.DocumentInfo.HasXmpMetadata);
        Assert.False(report.DocumentInfo.HasCatalogUri);
        Assert.Empty(report.ReadBlockers);
        Assert.Empty(report.RewriteBlockers);
        Assert.DoesNotContain("PDF output intents are not supported for rewriting by OfficeIMO.Pdf yet.", report.Diagnostics);
    }

    [Fact]
    public void Preflight_AllowsComplexOutputIntentPdfReadButBlocksRewrite() {
        PdfDocumentPreflight report = PdfInspector.Preflight(BuildComplexOutputIntentPdf());

        Assert.True(report.CanRead);
        Assert.False(report.CanRewrite);
        Assert.True(report.Probe.HasOutputIntents);
        Assert.NotNull(report.DocumentInfo);
        Assert.True(report.DocumentInfo!.HasOutputIntents);
        Assert.Empty(report.ReadBlockers);
        Assert.Contains("PDF output intents are not supported for rewriting by OfficeIMO.Pdf yet.", report.Diagnostics);
        AssertRewriteBlocker(report, PdfRewriteBlockerKind.OutputIntents, "PDF output intents are not supported for rewriting by OfficeIMO.Pdf yet.");
    }


}
