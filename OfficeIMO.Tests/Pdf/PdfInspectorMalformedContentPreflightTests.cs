using System;
using System.IO;
using System.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfInspectorTests {
    [Fact]
    public void Preflight_ReportsInvalidHeaderWithoutParserException() {
        PdfDocumentPreflight report = PdfInspector.Preflight(System.Text.Encoding.ASCII.GetBytes("not a pdf"));

        Assert.False(report.CanRead);
        Assert.False(report.CanRewrite);
        Assert.False(report.CanExtractText);
        Assert.False(report.CanExtractImages);
        Assert.False(report.CanReadLogicalObjects);
        Assert.False(report.CanManipulatePages);
        Assert.False(report.CanFillSimpleFormFields);
        Assert.False(report.CanFlattenSimpleFormFields);
        Assert.False(report.CanFillAndFlattenSimpleFormFields);
        Assert.Null(report.DocumentInfo);
        Assert.Null(report.Probe.HeaderVersion);
        Assert.Contains("PDF header was not found.", report.Diagnostics);
        AssertReadBlocker(report, PdfReadBlockerKind.MissingHeader, "PDF header was not found.");
        Assert.Empty(report.RewriteBlockers);
    }

    [Fact]
    public void Preflight_ReportsNoPagesWithReadBlocker() {
        PdfDocumentPreflight report = PdfInspector.Preflight(BuildNoPagesPdf());

        Assert.False(report.CanRead);
        Assert.False(report.CanRewrite);
        Assert.False(report.CanExtractText);
        Assert.False(report.CanExtractImages);
        Assert.False(report.CanReadLogicalObjects);
        Assert.False(report.CanManipulatePages);
        Assert.False(report.CanFillSimpleFormFields);
        Assert.False(report.CanFlattenSimpleFormFields);
        Assert.False(report.CanFillAndFlattenSimpleFormFields);
        Assert.NotNull(report.DocumentInfo);
        Assert.Empty(report.DocumentInfo!.Pages);
        Assert.Contains("No PDF pages were discovered.", report.Diagnostics);
        AssertReadBlocker(report, PdfReadBlockerKind.NoPages, "No PDF pages were discovered.");
        Assert.Empty(report.RewriteBlockers);
    }

    [Fact]
    public void Preflight_ReportsUnsupportedContentStreamFiltersWithReadBlocker() {
        PdfDocumentPreflight report = PdfInspector.Preflight(BuildUnsupportedContentStreamFilterPdf());

        Assert.False(report.CanRead);
        Assert.False(report.CanRewrite);
        Assert.False(report.CanExtractText);
        Assert.True(report.CanExtractImages);
        Assert.False(report.CanReadLogicalObjects);
        Assert.False(report.Can(PdfPreflightCapability.ExtractText));
        Assert.True(report.Can(PdfPreflightCapability.ExtractImages));
        Assert.False(report.Can(PdfPreflightCapability.ReadLogicalObjects));
        Assert.Contains(
            "PDF page content streams use unsupported filter(s): DCTDecode.",
            report.GetCapabilityDiagnostics(PdfPreflightCapability.ExtractText));
        Assert.Empty(report.GetCapabilityDiagnostics(PdfPreflightCapability.ExtractImages));
        Assert.Contains(
            "PDF page content streams use unsupported filter(s): DCTDecode.",
            report.GetCapabilityDiagnostics(PdfPreflightCapability.ReadLogicalObjects));
        Assert.NotNull(report.DocumentInfo);
        Assert.Equal(1, report.DocumentInfo!.PageCount);
        AssertReadBlocker(
            report,
            PdfReadBlockerKind.UnsupportedContentStreamFilter,
            "PDF page content streams use unsupported filter(s): DCTDecode.");
    }

    [Fact]
    public void Preflight_ReportsUnsupportedFormXObjectStreamFiltersWithReadBlocker() {
        PdfDocumentPreflight report = PdfInspector.Preflight(BuildUnsupportedFormXObjectStreamFilterPdf());

        Assert.False(report.CanRead);
        Assert.False(report.CanRewrite);
        Assert.NotNull(report.DocumentInfo);
        Assert.Equal(1, report.DocumentInfo!.PageCount);
        AssertReadBlocker(
            report,
            PdfReadBlockerKind.UnsupportedContentStreamFilter,
            "PDF page content streams use unsupported filter(s): DCTDecode.");
    }

    [Fact]
    public void Preflight_ReportsUnsupportedFormXObjectFiltersAcrossSplitContentStreams() {
        PdfDocumentPreflight report = PdfInspector.Preflight(BuildUnsupportedFormXObjectFilterSplitAcrossContentStreamsPdf());

        Assert.False(report.CanRead);
        Assert.False(report.CanRewrite);
        Assert.NotNull(report.DocumentInfo);
        Assert.Equal(1, report.DocumentInfo!.PageCount);
        AssertReadBlocker(
            report,
            PdfReadBlockerKind.UnsupportedContentStreamFilter,
            "PDF page content streams use unsupported filter(s): DCTDecode.");
    }

    [Fact]
    public void Preflight_ReportsUnsupportedFormXObjectFiltersWhenNameTokenIsSplitAcrossContentStreams() {
        PdfDocumentPreflight report = PdfInspector.Preflight(BuildUnsupportedFormXObjectFilterSplitAcrossContentStreamsPdf("q\n/Fm", "1 Do\nQ"));

        Assert.False(report.CanRead);
        Assert.False(report.CanRewrite);
        Assert.NotNull(report.DocumentInfo);
        Assert.Equal(1, report.DocumentInfo!.PageCount);
        AssertReadBlocker(
            report,
            PdfReadBlockerKind.UnsupportedContentStreamFilter,
            "PDF page content streams use unsupported filter(s): DCTDecode.");
    }

    [Fact]
    public void Preflight_BlocksWrongGenerationRewriteReferences() {
        PdfDocumentPreflight report = PdfInspector.Preflight(BuildWrongGenerationContentReferencePdf());

        Assert.True(report.CanRead);
        Assert.False(report.CanRewrite);
        Assert.NotNull(report.DocumentInfo);
        Assert.Equal(1, report.DocumentInfo!.PageCount);
        AssertRewriteBlocker(
            report,
            PdfRewriteBlockerKind.InvalidObjectReferences,
            "PDF object graph is not safe for rewriting by OfficeIMO.Pdf yet: PDF object 4 1 R was referenced, but the active object generation is 0.");
    }

    [Fact]
    public void Preflight_ReadsFromPathAndStream() {
        byte[] bytes = BuildTwoPagePdf();
        string path = Path.Combine(Path.GetTempPath(), "officeimo-pdf-preflight-" + Guid.NewGuid().ToString("N") + ".pdf");

        try {
            File.WriteAllBytes(path, bytes);

            PdfDocumentPreflight fromPath = PdfInspector.Preflight(path);
            using var stream = new MemoryStream(bytes);
            PdfDocumentPreflight fromStream = PdfInspector.Preflight(stream);

            Assert.True(fromPath.CanRewrite);
            Assert.True(fromStream.CanRewrite);
            Assert.Empty(fromPath.ReadBlockers);
            Assert.Empty(fromStream.ReadBlockers);
            Assert.Empty(fromPath.RewriteBlockers);
            Assert.Empty(fromStream.RewriteBlockers);
            Assert.Equal(2, fromPath.DocumentInfo!.PageCount);
            Assert.Equal(fromPath.DocumentInfo.PageCount, fromStream.DocumentInfo!.PageCount);
        } finally {
            if (File.Exists(path)) {
                File.Delete(path);
            }
        }
    }

    [Fact]
    public void Preflight_RejectsInvalidInputs() {
        Assert.Throws<ArgumentNullException>(() => PdfInspector.Preflight((byte[])null!));
        Assert.Throws<ArgumentNullException>(() => PdfInspector.Preflight((string)null!));
        Assert.Throws<ArgumentException>(() => PdfInspector.Preflight(" "));
        Assert.Throws<ArgumentNullException>(() => PdfInspector.Preflight((Stream)null!));
        Assert.Throws<ArgumentException>(() => PdfInspector.Preflight(new WriteOnlyStream()));

        PdfDocumentPreflight report = PdfInspector.Preflight(BuildTwoPagePdf());
        Assert.Throws<ArgumentOutOfRangeException>(() => report.Can((PdfPreflightCapability)999));
        Assert.Throws<ArgumentOutOfRangeException>(() => report.GetCapabilityDiagnostics((PdfPreflightCapability)999));
    }


}
