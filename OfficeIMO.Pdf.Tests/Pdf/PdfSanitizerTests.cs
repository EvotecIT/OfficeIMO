using System.Text;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfSanitizerTests {
    [Fact]
    public void Sanitize_RemovesActiveContentUnsafeUrisAndRichMediaButPreservesSafeLinks() {
        byte[] source = BuildActiveContentPdf();

        PdfSanitizationResult result = PdfSanitizer.Sanitize(source);
        byte[] sanitized = result.ToBytes();
        PdfDocumentInfo info = PdfInspector.Inspect(sanitized);

        Assert.True(result.IsSanitized);
        Assert.Equal(PdfMutationExecutionMode.FullRewrite, result.MutationPlan.ExecutionMode);
        Assert.Contains(PdfMutationProof.SanitizationReadback, result.MutationPlan.RequiredProofs);
        Assert.Contains(result.RemovedFindings, finding => finding.Kind == PdfSanitizationFindingKind.ActiveAction && finding.Detail == "JavaScript");
        Assert.Contains(result.RemovedFindings, finding => finding.Kind == PdfSanitizationFindingKind.ActiveAction && finding.Detail == "Launch");
        Assert.Contains(result.RemovedFindings, finding => finding.Kind == PdfSanitizationFindingKind.ActiveAction && finding.Detail == "SubmitForm");
        Assert.Contains(result.RemovedFindings, finding => finding.Kind == PdfSanitizationFindingKind.ActiveAction && finding.Detail == "GoToR");
        Assert.Contains(result.RemovedFindings, finding => finding.Kind == PdfSanitizationFindingKind.ActiveAction && finding.Detail == "GoToE");
        Assert.Contains(result.RemovedFindings, finding => finding.Kind == PdfSanitizationFindingKind.ActiveAction && finding.Detail == "ImportData");
        Assert.Contains(result.RemovedFindings, finding => finding.Kind == PdfSanitizationFindingKind.UnsafeUri && finding.Detail == "javascript:alert('unsafe')");
        Assert.Contains(result.RemovedFindings, finding => finding.Kind == PdfSanitizationFindingKind.RichMedia && finding.Detail == "RichMedia");
        Assert.Empty(result.RemainingFindings);
        Assert.False(info.HasActiveContent);
        Assert.Empty(info.CatalogActions);
        Assert.Empty(info.Pages[0].PageActions);
        Assert.Single(info.GetLinkAnnotationsByUri("https://example.com/safe"));
        Assert.Contains(info.Annotations, annotation => annotation.Subtype == "Text" && annotation.Contents == "keep me");
        Assert.Empty(PdfSanitizer.Analyze(sanitized));
        string raw = PdfEncoding.Latin1GetString(sanitized);
        Assert.DoesNotContain("app.alert", raw, StringComparison.Ordinal);
        Assert.DoesNotContain("tool.exe", raw, StringComparison.Ordinal);
    }

    [Fact]
    public void Sanitize_QuarantinesAttachmentsAndRemovesAllPayloadReferences() {
        byte[] payload = Encoding.UTF8.GetBytes("quarantined payload");
        var options = new PdfOptions().AddEmbeddedFile(
            "payload.txt",
            payload,
            "text/plain",
            PdfAssociatedFileRelationship.Data,
            "Sanitizer test payload");
        byte[] source = PdfDocument.Create(options)
            .Paragraph(paragraph => paragraph.Text("Attachment quarantine"))
            .ToBytes();
        var policy = new PdfSanitizationOptions {
            EmbeddedFiles = PdfEmbeddedFileSanitizationMode.Quarantine
        };

        PdfSanitizationResult result = PdfDocument.Open(source).Sanitize(policy);
        PdfDocumentInfo info = result.ToDocument().Inspect();

        PdfExtractedAttachment attachment = Assert.Single(result.QuarantinedAttachments);
        Assert.Equal("payload.txt", attachment.FileName);
        Assert.Equal(payload, attachment.Bytes);
        Assert.Empty(info.Attachments);
        Assert.False(info.HasEmbeddedFiles);
        Assert.Contains(result.RemovedFindings, finding => finding.Kind == PdfSanitizationFindingKind.EmbeddedFile);
        Assert.Empty(PdfAttachmentExtractor.ExtractAttachments(result.ToBytes()));
    }

    [Fact]
    public void Sanitize_ExplicitActionAllowListCanPreserveReviewedJavaScript() {
        byte[] source = BuildActiveContentPdf();
        var policy = new PdfSanitizationOptions();
        policy.AllowedActionTypes.Add("JavaScript");

        PdfSanitizationResult result = PdfSanitizer.Sanitize(source, policy);
        PdfDocumentInfo info = PdfInspector.Inspect(result.ToBytes());

        Assert.True(result.IsSanitized);
        Assert.Contains(info.CatalogActions, action => action.ActionType == "JavaScript");
        Assert.True(result.PreservationReport.IsPreserved);
        Assert.Contains(result.PreservationReport.Original.CatalogActions, action => action.ActionType == "JavaScript");
        Assert.Contains(result.PreservationReport.Rewritten.CatalogActions, action => action.ActionType == "JavaScript");
        Assert.DoesNotContain(result.RemovedFindings, finding => finding.Detail == "JavaScript");
        Assert.Contains(result.RemovedFindings, finding => finding.Detail == "Launch");
        Assert.Contains(result.RemovedFindings, finding => finding.Detail == "SubmitForm");
    }

    [Fact]
    public void Sanitize_PreservationIncludesSafeUriActionsAndOpeningDestinations() {
        byte[] source = BuildSafeViewerActionPdf();

        PdfSanitizationResult result = PdfSanitizer.Sanitize(source);
        PdfDocumentInfo info = result.ToDocument().Inspect();

        Assert.True(result.PreservationReport.IsPreserved);
        Assert.NotNull(info.OpenAction);
        Assert.Equal("Destination", info.OpenAction!.ActionType);
        Assert.Contains(info.Pages[0].PageActions, action => action.ActionType == "URI");
        Assert.Contains(info.Pages[0].PageActions, action => action.ActionType == "GoTo");
        PdfCatalogAction catalogUri = Assert.Single(info.CatalogActions, static action => action.ActionType == "URI");
        Assert.Equal("https://example.com/catalog", catalogUri.Uri);
    }

    [Fact]
    public void SanitizerPreservationFilterIncludesEveryPolicyRetainedActionType() {
        PdfDocumentInfo info = PdfInspector.Inspect(BuildSafeViewerActionPdf());
        var actionTypes = new HashSet<string>(StringComparer.Ordinal);

        PdfSanitizer.AddPolicyRetainedActionTypes(info, new PdfSanitizationOptions(), actionTypes);

        Assert.Contains("Destination", actionTypes);
        Assert.Contains("GoTo", actionTypes);
        Assert.Contains("URI", actionTypes);
        Assert.DoesNotContain("JavaScript", actionTypes);
    }

    [Fact]
    public void Sanitize_PreservesSafeUriActionsWhenUnsafeUriActionsAlsoExist() {
        byte[] source = BuildMixedPageUriActionPdf();

        PdfSanitizationResult result = PdfSanitizer.Sanitize(source);
        PdfDocumentInfo info = result.ToDocument().Inspect();

        PdfPageAction action = Assert.Single(info.Pages[0].PageActions);
        Assert.Equal("https://example.com/safe", action.Uri);
        Assert.True(result.PreservationReport.IsPreserved);
        Assert.Contains(result.RemovedFindings, finding => finding.Kind == PdfSanitizationFindingKind.UnsafeUri);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Sanitize_PreservesNextActionsWhenRemovedSiblingsShiftTheirIndices(bool catalogAction) {
        byte[] source = BuildMixedNextActionPdf(catalogAction);

        PdfSanitizationResult result = PdfSanitizer.Sanitize(source);
        PdfDocumentInfo info = result.ToDocument().Inspect();

        Assert.True(result.PreservationReport.IsPreserved);
        Assert.Contains(result.RemovedFindings, static finding => finding.Detail == "JavaScript");
        if (catalogAction) {
            Assert.Equal(2, info.CatalogActions.Count(static action => action.ActionType == "URI"));
        } else {
            Assert.Equal(2, info.Pages[0].PageActions.Count(static action => action.ActionType == "URI"));
        }
    }

    [Fact]
    public void Sanitize_ReusesCustomReadLimitsForItsRewrittenArtifact() {
        byte[] source = BuildActiveContentPdf();
        var readOptions = new PdfReadOptions {
            Limits = new PdfReadLimits {
                MaxInputBytes = source.LongLength,
                MaxJavaScripts = PdfReadLimits.DefaultMaxJavaScripts + 17
            }
        };

        PdfSanitizationResult result = PdfDocument.Open(source, readOptions).Sanitize();
        PdfDocument reopened = result.ToDocument();

        Assert.True(result.IsSanitized);
        Assert.True(result.PreservationReport.IsPreserved);
        Assert.Empty(result.RemainingFindings);
        Assert.Equal(PdfReadLimits.DefaultMaxJavaScripts + 17, reopened.ReadOptions.Limits.MaxJavaScripts);
        Assert.True(reopened.ReadOptions.Limits.MaxInputBytes >= result.ToBytes().LongLength);
        Assert.Empty(PdfSanitizer.Analyze(reopened.ToBytes()));
    }

    [Theory]
    [InlineData("RichMedia")]
    [InlineData("Movie")]
    [InlineData("Sound")]
    [InlineData("Screen")]
    [InlineData("3D")]
    [InlineData("FileAttachment")]
    public void Sanitize_ExcludesEveryPolicyRemovedRichAnnotationFromPreservation(string subtype) {
        byte[] source = BuildSingleAnnotationPdf(subtype);

        PdfSanitizationResult result = PdfSanitizer.Sanitize(source);

        Assert.True(result.PreservationReport.IsPreserved);
        Assert.Empty(PdfInspector.Inspect(result.ToBytes()).Annotations);
        Assert.Contains(result.RemovedFindings, finding => finding.Kind == PdfSanitizationFindingKind.RichMedia && finding.Detail == subtype);
    }

    [Fact]
    public void RewritePreservation_CanCompareOnlyAllowlistedActionTypes() {
        byte[] source = BuildActiveContentPdf();
        var policy = new PdfSanitizationOptions();
        policy.AllowedActionTypes.Add("JavaScript");
        byte[] sanitized = PdfSanitizer.Sanitize(source, policy).ToBytes();
        var options = new PdfRewritePreservationOptions {
            PreserveCatalogActions = true,
            PreservePageActions = true,
            PreserveOpenAction = true,
            PreserveLinkAnnotations = false,
            PreserveAnnotations = false,
            PreserveEmbeddedFiles = false,
            PreserveRevisionStructure = false
        };
        options.PreservedActionTypes.Add("JavaScript");

        PdfRewritePreservationReport report = PdfRewritePreservation.Assess(source, sanitized, options);

        Assert.True(report.IsPreserved, string.Join(Environment.NewLine, report.Issues.Select(static issue => issue.ToString())));
    }

    [Fact]
    public void Sanitize_QuarantinesPageAssociatedFilePayloads() {
        byte[] source = PdfAssociatedFileTestSupport.BuildPageAssociatedFilePdf();
        var policy = new PdfSanitizationOptions { EmbeddedFiles = PdfEmbeddedFileSanitizationMode.Quarantine };

        PdfSanitizationResult result = PdfSanitizer.Sanitize(source, policy);

        PdfExtractedAttachment attachment = Assert.Single(result.QuarantinedAttachments);
        Assert.Equal("page.txt", attachment.FileName);
        Assert.Equal(PdfAssociatedFileTestSupport.Payload, Encoding.ASCII.GetString(attachment.Bytes));
        Assert.Empty(PdfAttachmentExtractor.ExtractAttachments(result.ToBytes()));
        Assert.DoesNotContain(PdfAssociatedFileTestSupport.Payload, Encoding.ASCII.GetString(result.ToBytes()), StringComparison.Ordinal);
    }

    private static byte[] BuildActiveContentPdf() {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.7",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R /Names << /JavaScript << /Names [(Open) 6 0 R] >> >> /AA << /WC 12 0 R /WS 13 0 R /WP 14 0 R >> >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 320 220] /Contents 4 0 R /Annots [5 0 R 9 0 R 10 0 R 11 0 R 15 0 R] /AA << /O 7 0 R >> >>",
            "endobj",
            "4 0 obj",
            "<< /Length 0 >>",
            "stream",
            string.Empty,
            "endstream",
            "endobj",
            "5 0 obj",
            "<< /Type /Annot /Subtype /Link /Rect [20 160 120 180] /A << /S /Launch /F (tool.exe) >> /AA << /E 8 0 R >> >>",
            "endobj",
            "6 0 obj",
            "<< /S /JavaScript /JS (app.alert('catalog')) >>",
            "endobj",
            "7 0 obj",
            "<< /S /JavaScript /JS (app.alert('page')) >>",
            "endobj",
            "8 0 obj",
            "<< /S /SubmitForm /F (https://example.com/submit) >>",
            "endobj",
            "9 0 obj",
            "<< /Type /Annot /Subtype /Link /Rect [20 120 180 140] /A << /S /URI /URI (https://example.com/safe) >> >>",
            "endobj",
            "10 0 obj",
            "<< /Type /Annot /Subtype /Link /Rect [20 80 180 100] /A << /S /URI /URI (javascript:alert('unsafe')) >> >>",
            "endobj",
            "11 0 obj",
            "<< /Type /Annot /Subtype /RichMedia /Rect [20 20 180 60] /RichMediaContent << >> >>",
            "endobj",
            "12 0 obj",
            "<< /S /GoToR /F (remote.pdf) /D [0 /Fit] >>",
            "endobj",
            "13 0 obj",
            "<< /S /ImportData /F (form-data.fdf) >>",
            "endobj",
            "14 0 obj",
            "<< /S /GoToE /F << /F (embedded.pdf) >> /D [0 /Fit] >>",
            "endobj",
            "15 0 obj",
            "<< /Type /Annot /Subtype /Text /Rect [200 20 220 40] /Contents (keep me) >>",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 16 >>",
            "%%EOF"
        }) + "\n";

        return Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildMixedNextActionPdf(bool catalogAction) {
        string action = "<< /S /URI /URI (https://example.com/one) /Next [<< /S /JavaScript /JS (app.alert\\('remove'\\);) >> << /S /URI /URI (https://example.com/two) >>] >>";
        string catalogEntry = catalogAction ? " /AA << /WC " + action + " >>" : string.Empty;
        string pageEntry = catalogAction ? string.Empty : " /AA << /O " + action + " >>";
        return Encoding.ASCII.GetBytes(string.Join("\n", new[] {
            "%PDF-1.7",
            "1 0 obj", "<< /Type /Catalog /Pages 2 0 R" + catalogEntry + " >>", "endobj",
            "2 0 obj", "<< /Type /Pages /Count 1 /Kids [3 0 R] >>", "endobj",
            "3 0 obj", "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 4 0 R" + pageEntry + " >>", "endobj",
            "4 0 obj", "<< /Length 0 >>", "stream", string.Empty, "endstream", "endobj",
            "trailer", "<< /Root 1 0 R /Size 5 >>", "%%EOF"
        }));
    }

    private static byte[] BuildSingleAnnotationPdf(string subtype) {
        string annotation = "<< /Type /Annot /Subtype /" + subtype + " /Rect [20 20 180 60] >>";
        string pdf = "%PDF-1.7\n" +
            "1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj\n" +
            "2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] >>\nendobj\n" +
            "3 0 obj\n<< /Type /Page /Parent 2 0 R /MediaBox [0 0 320 220] /Contents 4 0 R /Annots [5 0 R] >>\nendobj\n" +
            "4 0 obj\n<< /Length 0 >>\nstream\n\nendstream\nendobj\n" +
            "5 0 obj\n" + annotation + "\nendobj\n" +
            "trailer\n<< /Root 1 0 R /Size 6 >>\n%%EOF\n";
        return Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildSafeViewerActionPdf() {
        string pdf = "%PDF-1.7\n" +
            "1 0 obj\n<< /Type /Catalog /Pages 2 0 R /OpenAction [3 0 R /Fit] /AA << /WC << /S /URI /URI (https://example.com/catalog) >> >> >>\nendobj\n" +
            "2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] >>\nendobj\n" +
            "3 0 obj\n<< /Type /Page /Parent 2 0 R /MediaBox [0 0 320 220] /Contents 4 0 R /AA << /O << /S /URI /URI (https://example.com/safe) >> /C << /S /GoTo /D [3 0 R /Fit] >> >> >>\nendobj\n" +
            "4 0 obj\n<< /Length 0 >>\nstream\n\nendstream\nendobj\n" +
            "trailer\n<< /Root 1 0 R /Size 5 >>\n%%EOF\n";
        return Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildMixedPageUriActionPdf() {
        string pdf = "%PDF-1.7\n" +
            "1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj\n" +
            "2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] >>\nendobj\n" +
            "3 0 obj\n<< /Type /Page /Parent 2 0 R /MediaBox [0 0 320 220] /Contents 4 0 R /AA << /O << /S /URI /URI (https://example.com/safe) >> /C << /S /URI /URI (javascript:unsafe) >> >> >>\nendobj\n" +
            "4 0 obj\n<< /Length 0 >>\nstream\n\nendstream\nendobj\n" +
            "trailer\n<< /Root 1 0 R /Size 5 >>\n%%EOF\n";
        return Encoding.ASCII.GetBytes(pdf);
    }
}
