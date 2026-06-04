using System;
using System.IO;
using System.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfInspectorTests {
    [Fact]
    public void Preflight_AllowsSimpleEmbeddedFilePdfReadAndRewrite() {
        PdfDocumentPreflight report = PdfInspector.Preflight(BuildEmbeddedFilePdf());

        Assert.True(report.CanRead);
        Assert.True(report.CanRewrite);
        Assert.True(report.Probe.HasEmbeddedFiles);
        Assert.True(report.Probe.HasCatalogNameTrees);
        Assert.False(report.Probe.HasOptionalContent);
        Assert.NotNull(report.DocumentInfo);
        Assert.True(report.DocumentInfo!.HasEmbeddedFiles);
        Assert.True(report.DocumentInfo.HasCatalogNameTrees);
        Assert.False(report.DocumentInfo.HasOptionalContent);
        Assert.Empty(report.ReadBlockers);
        Assert.Empty(report.RewriteBlockers);
        Assert.DoesNotContain("PDF embedded files are not supported for rewriting by OfficeIMO.Pdf yet.", report.Diagnostics);
    }

    [Fact]
    public void Preflight_AllowsSimpleAssociatedFilePdfReadAndRewrite() {
        PdfDocumentPreflight report = PdfInspector.Preflight(BuildAssociatedFilePdf());

        Assert.True(report.CanRead);
        Assert.True(report.CanRewrite);
        Assert.True(report.Probe.HasEmbeddedFiles);
        Assert.False(report.Probe.HasCatalogNameTrees);
        Assert.False(report.Probe.HasOptionalContent);
        Assert.NotNull(report.DocumentInfo);
        Assert.True(report.DocumentInfo!.HasEmbeddedFiles);
        Assert.False(report.DocumentInfo.HasCatalogNameTrees);
        Assert.False(report.DocumentInfo.HasOptionalContent);
        Assert.Empty(report.ReadBlockers);
        Assert.Empty(report.RewriteBlockers);
        Assert.DoesNotContain("PDF embedded files are not supported for rewriting by OfficeIMO.Pdf yet.", report.Diagnostics);
    }

    [Fact]
    public void Preflight_AllowsCombinedDestinationAndEmbeddedFileNameTreesReadAndRewrite() {
        PdfDocumentPreflight report = PdfInspector.Preflight(BuildCombinedDestinationAndEmbeddedFileNameTreePdf());

        Assert.True(report.CanRead);
        Assert.True(report.CanRewrite);
        Assert.True(report.Probe.HasNamedDestinations);
        Assert.True(report.Probe.HasEmbeddedFiles);
        Assert.True(report.Probe.HasCatalogNameTrees);
        Assert.NotNull(report.DocumentInfo);
        Assert.True(report.DocumentInfo!.HasNamedDestinations);
        Assert.True(report.DocumentInfo.HasEmbeddedFiles);
        Assert.True(report.DocumentInfo.HasCatalogNameTrees);
        Assert.Empty(report.ReadBlockers);
        Assert.Empty(report.RewriteBlockers);
        Assert.DoesNotContain("PDF named destinations are not supported for rewriting by OfficeIMO.Pdf yet.", report.Diagnostics);
        Assert.DoesNotContain("PDF embedded files are not supported for rewriting by OfficeIMO.Pdf yet.", report.Diagnostics);
    }

    [Fact]
    public void Preflight_AllowsUnsupportedCatalogNameTreePdfReadButBlocksRewrite() {
        PdfDocumentPreflight report = PdfInspector.Preflight(BuildUnsupportedCatalogNameTreePdf());

        Assert.True(report.CanRead);
        Assert.False(report.CanRewrite);
        Assert.True(report.Probe.HasCatalogNameTrees);
        Assert.False(report.Probe.HasNamedDestinations);
        Assert.False(report.Probe.HasEmbeddedFiles);
        Assert.NotNull(report.DocumentInfo);
        Assert.True(report.DocumentInfo!.HasCatalogNameTrees);
        Assert.False(report.DocumentInfo.HasNamedDestinations);
        Assert.False(report.DocumentInfo.HasEmbeddedFiles);
        Assert.Empty(report.ReadBlockers);
        Assert.Contains("PDF catalog name trees are not supported for rewriting by OfficeIMO.Pdf yet.", report.Diagnostics);
        AssertRewriteBlocker(report, PdfRewriteBlockerKind.CatalogNameTrees, "PDF catalog name trees are not supported for rewriting by OfficeIMO.Pdf yet.");
    }

    [Fact]
    public void Preflight_AllowsComplexEmbeddedFilePdfReadButBlocksRewrite() {
        PdfDocumentPreflight report = PdfInspector.Preflight(BuildComplexEmbeddedFilePdf());

        Assert.True(report.CanRead);
        Assert.False(report.CanRewrite);
        Assert.True(report.Probe.HasEmbeddedFiles);
        Assert.True(report.Probe.HasCatalogNameTrees);
        Assert.NotNull(report.DocumentInfo);
        Assert.True(report.DocumentInfo!.HasEmbeddedFiles);
        Assert.True(report.DocumentInfo.HasCatalogNameTrees);
        Assert.Empty(report.ReadBlockers);
        Assert.Contains("PDF embedded files are not supported for rewriting by OfficeIMO.Pdf yet.", report.Diagnostics);
        AssertRewriteBlocker(report, PdfRewriteBlockerKind.EmbeddedFiles, "PDF embedded files are not supported for rewriting by OfficeIMO.Pdf yet.");
    }

    [Fact]
    public void Preflight_AllowsComplexAssociatedFilePdfReadButBlocksRewrite() {
        PdfDocumentPreflight report = PdfInspector.Preflight(BuildComplexAssociatedFilePdf());

        Assert.True(report.CanRead);
        Assert.False(report.CanRewrite);
        Assert.True(report.Probe.HasEmbeddedFiles);
        Assert.False(report.Probe.HasCatalogNameTrees);
        Assert.NotNull(report.DocumentInfo);
        Assert.True(report.DocumentInfo!.HasEmbeddedFiles);
        Assert.False(report.DocumentInfo.HasCatalogNameTrees);
        Assert.Empty(report.ReadBlockers);
        Assert.Contains("PDF embedded files are not supported for rewriting by OfficeIMO.Pdf yet.", report.Diagnostics);
        AssertRewriteBlocker(report, PdfRewriteBlockerKind.EmbeddedFiles, "PDF embedded files are not supported for rewriting by OfficeIMO.Pdf yet.");
    }

    [Fact]
    public void Preflight_AllowsSimpleOptionalContentPdfReadAndRewrite() {
        PdfDocumentPreflight report = PdfInspector.Preflight(BuildOptionalContentPdf());

        Assert.True(report.CanRead);
        Assert.True(report.CanRewrite);
        Assert.True(report.Probe.HasOptionalContent);
        Assert.False(report.Probe.HasEmbeddedFiles);
        Assert.NotNull(report.DocumentInfo);
        Assert.True(report.DocumentInfo!.HasOptionalContent);
        Assert.False(report.DocumentInfo.HasEmbeddedFiles);
        Assert.Empty(report.ReadBlockers);
        Assert.Empty(report.RewriteBlockers);
        Assert.DoesNotContain("PDF optional content layers are not supported for rewriting by OfficeIMO.Pdf yet.", report.Diagnostics);
    }

    [Fact]
    public void Preflight_AllowsComplexOptionalContentPdfReadButBlocksRewrite() {
        PdfDocumentPreflight report = PdfInspector.Preflight(BuildComplexOptionalContentPdf());

        Assert.True(report.CanRead);
        Assert.False(report.CanRewrite);
        Assert.True(report.Probe.HasOptionalContent);
        Assert.NotNull(report.DocumentInfo);
        Assert.True(report.DocumentInfo!.HasOptionalContent);
        Assert.Empty(report.ReadBlockers);
        Assert.Contains("PDF optional content layers are not supported for rewriting by OfficeIMO.Pdf yet.", report.Diagnostics);
        AssertRewriteBlocker(report, PdfRewriteBlockerKind.OptionalContent, "PDF optional content layers are not supported for rewriting by OfficeIMO.Pdf yet.");
    }

    [Fact]
    public void Preflight_AllowsActiveContentPdfReadButBlocksRewrite() {
        PdfDocumentPreflight report = PdfInspector.Preflight(BuildActiveContentPdf());

        Assert.True(report.CanRead);
        Assert.False(report.CanRewrite);
        Assert.True(report.CanExtractText);
        Assert.True(report.CanExtractImages);
        Assert.True(report.CanReadLogicalObjects);
        Assert.False(report.CanManipulatePages);
        Assert.False(report.CanFillSimpleFormFields);
        Assert.False(report.CanFlattenSimpleFormFields);
        Assert.False(report.CanFillAndFlattenSimpleFormFields);
        Assert.False(report.Can(PdfPreflightCapability.FillSimpleFormFields));
        Assert.False(report.Can(PdfPreflightCapability.FlattenSimpleFormFields));
        Assert.Contains(
            "PDF active content is not supported for form filling or flattening by OfficeIMO.Pdf yet.",
            report.GetCapabilityDiagnostics(PdfPreflightCapability.FillSimpleFormFields));
        Assert.True(report.Probe.HasActiveContent);
        Assert.True(report.Probe.HasCatalogNameTrees);
        Assert.NotNull(report.DocumentInfo);
        Assert.True(report.DocumentInfo!.HasActiveContent);
        Assert.True(report.DocumentInfo.HasCatalogNameTrees);
        Assert.Empty(report.ReadBlockers);
        Assert.Contains("PDF active content is not supported for rewriting by OfficeIMO.Pdf yet.", report.Diagnostics);
        AssertRewriteBlocker(report, PdfRewriteBlockerKind.ActiveContent, "PDF active content is not supported for rewriting by OfficeIMO.Pdf yet.");
    }


}
