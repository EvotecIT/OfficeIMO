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
    public void AppliedRedactionPlanVerificationFailsWhenResidualInspectionIsBlocked() {
        byte[] source = PdfDocument.Create().Paragraph(paragraph => paragraph.Text("Reviewed secret")).ToBytes();
        PdfRedactionPlan reviewedPlan = PdfDocument.Load(source).Redactions.Plan([
            new PdfRedactionArea(1, 0D, 0D, 600D, 800D, "reviewed")
        ]);

        PdfRedactionVerificationReport report = PdfRedactionVerification.VerifyAppliedPlan(
            BuildNoPagesPdf(),
            reviewedPlan,
            new PdfRedactionVerificationOptions());

        Assert.False(report.IsVerified);
        Assert.Contains(report.Issues, issue => issue.Feature == "RedactionPlanInspectionBlocked");
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
    public void PageTransfer_PreservesUnsupportedContentStreamsWithoutWeakeningGeneralPreflight() {
        byte[] source = BuildUnsupportedContentStreamFilterPdf();
        PdfDocument opened = PdfDocument.Load(source);

        byte[] extracted = opened.Pages.Extract(1).ToBytes();
        byte[] merged = PdfDocument.MergeBytes(new[] { source, BuildTwoPagePdf() }).ToBytes();
        PdfDocumentPreflight preflight = opened.Preflight();
        PdfMutationPortfolioReport portfolio = opened.AssessMutations(new[] {
            PdfMutationOperation.ExtractPages,
            PdfMutationOperation.MergeDocuments
        });
        PdfOperationResult<PdfDocument> extractionResult = opened.Pages.ExtractResult(PdfPageSelection.From(1));
        PdfOperationResult<PdfDocument> primaryMergeResult = opened.MergeWithResult(PdfDocument.Load(BuildTwoPagePdf()));
        PdfOperationResult<PdfDocument> incomingMergeResult = PdfDocument.Load(BuildTwoPagePdf()).MergeWithResult(opened);

        Assert.False(preflight.CanRead);
        Assert.True(preflight.CanManipulatePages);
        Assert.True(opened.PlanMutation(PdfMutationOperation.ExtractPages).CanExecute);
        Assert.True(opened.PlanMutation(PdfMutationOperation.MergeDocuments).CanExecute);
        Assert.True(portfolio.CanExecuteAll);
        Assert.True(extractionResult.Succeeded);
        Assert.True(primaryMergeResult.Succeeded);
        Assert.True(incomingMergeResult.Succeeded);
        Assert.False(PdfInspector.Preflight(extracted).CanRead);
        Assert.Single(PdfReadDocument.Open(extracted).Pages);
        Assert.Equal(3, PdfReadDocument.Open(merged).Pages.Count);
        Assert.Contains("/DCTDecode", PdfEncoding.Latin1GetString(extracted), StringComparison.Ordinal);
        Assert.Contains("/DCTDecode", PdfEncoding.Latin1GetString(merged), StringComparison.Ordinal);
        Assert.Equal(GetFilteredStreamData(source, "DCTDecode"), GetFilteredStreamData(extracted, "DCTDecode"));
        Assert.Equal(GetFilteredStreamData(source, "DCTDecode"), GetFilteredStreamData(merged, "DCTDecode"));
    }

    [Fact]
    public void PageTransfer_BlocksContextDependentCryptStreamsAcrossAllEntryPoints() {
        byte[] source = BuildUnsupportedContentStreamFilterPdf("/Crypt");
        PdfDocument opened = PdfDocument.Load(source);
        PdfDocumentPreflight preflight = opened.Preflight();

        Assert.False(preflight.CanRead);
        Assert.False(preflight.CanManipulatePages);
        Assert.True(preflight.HasReadBlocker(PdfReadBlockerKind.ContextDependentContentStreamFilter));
        Assert.False(opened.PlanMutation(PdfMutationOperation.ExtractPages).CanExecute);
        Assert.False(opened.PlanMutation(PdfMutationOperation.MergeDocuments).CanExecute);
        Assert.False(opened.Pages.ExtractResult(PdfPageSelection.From(1)).CanAttempt);
        Assert.False(opened.MergeWithResult(PdfDocument.Load(BuildTwoPagePdf())).CanAttempt);
        Assert.Throws<PdfMutationBlockedException>(() => opened.Pages.Extract(1));
        Assert.Throws<PdfMutationBlockedException>(() => PdfDocument.MergeBytes(new[] { source, BuildTwoPagePdf() }));
    }

    [Fact]
    public void PageTransfer_BlocksContextDependentCryptOutsideTheSelectedPageGraph() {
        string source = PdfEncoding.Latin1GetString(BuildUnsupportedContentStreamFilterPdf());
        byte[] withOrphanCryptStream = PdfEncoding.Latin1GetBytes(source.Replace(
            "trailer\n<< /Root 1 0 R /Size 5 >>",
            "5 0 obj\n<< /Length 4 /Filter /Crypt >>\nstream\ndata\nendstream\nendobj\ntrailer\n<< /Root 1 0 R /Size 6 >>"));
        PdfDocument opened = PdfDocument.Load(withOrphanCryptStream);
        PdfDocumentPreflight preflight = opened.Preflight();

        Assert.True(preflight.HasReadBlocker(PdfReadBlockerKind.UnsupportedContentStreamFilter));
        Assert.True(preflight.HasReadBlocker(PdfReadBlockerKind.ContextDependentContentStreamFilter));
        Assert.False(preflight.CanManipulatePages);
        Assert.False(opened.PlanMutation(PdfMutationOperation.ExtractPages).CanExecute);
        Assert.False(opened.PlanMutation(PdfMutationOperation.MergeDocuments).CanExecute);
        Assert.Throws<PdfMutationBlockedException>(() => opened.Pages.Extract(1));
    }

    [Fact]
    public void BlockedPageTransfer_RebuildsCompletePreflightDiagnostics() {
        string source = PdfEncoding.Latin1GetString(BuildUnsupportedContentStreamFilterPdf());
        byte[] signedSource = PdfEncoding.Latin1GetBytes(source.Replace(
            "trailer\n<< /Root 1 0 R /Size 5 >>",
            "5 0 obj\n<< /Type /Sig /ByteRange [0 1 2 3] >>\nendobj\ntrailer\n<< /Root 1 0 R /Size 6 >>"));

        PdfMutationBlockedException exception = Assert.Throws<PdfMutationBlockedException>(() =>
            PdfDocument.Load(signedSource).Pages.Extract(1));

        Assert.True(exception.Plan.Preflight.HasReadBlocker(PdfReadBlockerKind.UnsupportedContentStreamFilter));
        Assert.True(exception.Plan.Preflight.HasRewriteBlocker(PdfRewriteBlockerKind.Signatures));
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

    private static byte[] GetFilteredStreamData(byte[] pdf, string filterName) {
        PdfReadDocument document = PdfReadDocument.Open(pdf);
        return document.Objects.Values
            .Select(static indirectObject => indirectObject.Value)
            .OfType<PdfStream>()
            .Single(stream => string.Equals(
                stream.Dictionary.Get<PdfName>("Filter")?.Name,
                filterName,
                StringComparison.Ordinal))
            .Data;
    }


}
