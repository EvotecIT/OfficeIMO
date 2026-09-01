using System.Text;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfSanitizerTests {
    [Fact]
    public void Sanitize_StopsWhileSerializingAtTheConfiguredOutputLimit() {
        byte[] source = BuildActiveContentPdf();

        InvalidOperationException exception = Assert.Throws<InvalidOperationException>(() =>
            PdfSanitizer.Sanitize(source, new PdfSanitizationOptions { MaximumOutputBytes = 128L }));

        Assert.Contains("while it was being serialized", exception.Message, StringComparison.Ordinal);
    }

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

        PdfSanitizationResult result = PdfDocument.Load(source).Sanitize(policy);
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
    public void Sanitize_PreservesCatalogActionsThatThePolicyRetains() {
        byte[] source = BuildSafeViewerActionPdf();
        var policy = new PdfSanitizationOptions();
        policy.AllowedActionTypes.Add("GoToR");

        PdfSanitizationResult result = PdfSanitizer.Sanitize(source, policy);
        PdfDocumentInfo info = result.ToDocument().Inspect();

        Assert.True(result.PreservationReport.IsPreserved);
        Assert.Contains(info.CatalogActions, static action => action.ActionType == "GoTo");
        Assert.Contains(info.CatalogActions, static action => action.ActionType == "GoToR");
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
    public void Sanitize_PromotesRetainedWidgetNextActionAndProvesItsPreservation() {
        byte[] source = BuildForbiddenWidgetRootWithRetainedNextPdf();

        PdfSanitizationResult result = PdfSanitizer.Sanitize(source);
        PdfFormWidget widget = Assert.Single(Assert.Single(result.ToDocument().Inspect().FormFields).Widgets);
        PdfFormWidgetAction action = Assert.Single(widget.Actions);

        Assert.Equal("GoTo", action.ActionType);
        Assert.True(result.PreservationReport.IsPreserved, result.PreservationReport.Summary);
        Assert.Contains(result.RemovedFindings, static finding => finding.Detail == "JavaScript");
    }

    [Fact]
    public void Sanitize_RemovesUnsafeWidgetUriWithoutFailingPreservation() {
        PdfSanitizationResult result = PdfSanitizer.Sanitize(BuildUnsafeWidgetUriPdf());

        PdfFormWidget widget = Assert.Single(Assert.Single(result.ToDocument().Inspect().FormFields).Widgets);
        Assert.Empty(widget.Actions);
        Assert.True(result.PreservationReport.IsPreserved, result.PreservationReport.Summary);
        Assert.Contains(result.RemovedFindings, static finding =>
            finding.Kind == PdfSanitizationFindingKind.UnsafeUri && finding.Detail == "javascript:unsafe");
    }

    [Fact]
    public void Reader_ExposesSafeWidgetUriTarget() {
        byte[] source = Encoding.ASCII.GetBytes(
            Encoding.ASCII.GetString(BuildUnsafeWidgetUriPdf())
                .Replace("javascript:unsafe", "https://example.com"));

        PdfFormWidgetAction action = Assert.Single(Assert.Single(Assert.Single(PdfDocument.Load(source).Inspect().FormFields).Widgets).Actions);

        Assert.Equal("URI", action.ActionType);
        Assert.Equal("https://example.com", action.Uri);
    }

    [Fact]
    public void Sanitize_PromotesAllowedDescendantFromForbiddenNextArrayEntry() {
        PdfSanitizationResult result = PdfSanitizer.Sanitize(BuildAllowedRootWithForbiddenNextDescendantPdf());

        PdfPageAction[] actions = result.ToDocument().Inspect().Pages[0].PageActions.ToArray();
        Assert.Equal(2, actions.Length);
        Assert.All(actions, static action => Assert.Equal("URI", action.ActionType));
        Assert.Contains(actions, static action => action.Uri == "https://example.com/root");
        Assert.Contains(actions, static action => action.Uri == "https://example.com/promoted");
        Assert.True(result.PreservationReport.IsPreserved, result.PreservationReport.Summary);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Sanitize_PromotesRetainedViewerNextActionAndProvesItsPreservation(bool catalogAction) {
        byte[] source = BuildForbiddenViewerRootWithRetainedNextPdf(catalogAction);

        PdfSanitizationResult result = PdfSanitizer.Sanitize(source);
        PdfDocumentInfo info = result.ToDocument().Inspect();
        string actionType = catalogAction
            ? Assert.Single(info.CatalogActions).ActionType
            : Assert.Single(info.Pages[0].PageActions).ActionType;

        Assert.Equal("GoTo", actionType);
        Assert.True(result.PreservationReport.IsPreserved, result.PreservationReport.Summary);
        Assert.Contains(result.RemovedFindings, static finding => finding.Detail == "JavaScript");
    }

    [Fact]
    public void Sanitize_PromotesRetainedOpenActionDescendantWithoutAFalsePreservationFailure() {
        PdfSanitizationResult result = PdfSanitizer.Sanitize(BuildForbiddenOpenActionWithRetainedNextPdf());
        PdfDocumentInfo info = result.ToDocument().Inspect();

        Assert.NotNull(info.OpenAction);
        Assert.Equal("GoTo", info.OpenAction!.ActionType);
        Assert.True(result.PreservationReport.IsPreserved, result.PreservationReport.Summary);
        Assert.Contains(result.RemovedFindings, static finding => finding.Detail == "JavaScript");
    }

    [Fact]
    public void Sanitize_BoundsSharedRetainedActionDagExpansion() {
        byte[] source = BuildSharedRetainedActionDagPdf(depth: 8);
        var readOptions = new PdfLoadOptions { Limits = new PdfReadLimits { MaxIndirectObjects = 16 } };

        PdfReadLimitException exception = Assert.Throws<PdfReadLimitException>(() =>
            PdfDocument.Load(source, readOptions).Sanitize());

        Assert.Equal(PdfReadLimitKind.IndirectObjects, exception.Kind);
        Assert.Equal(16, exception.Limit);
        Assert.Equal(17, exception.Actual);
    }

    [Fact]
    public void Sanitize_CountsEachRetainedNextActionOnce() {
        byte[] source = BuildLinearRetainedActionChainPdf(actionCount: 5);
        var readOptions = new PdfLoadOptions { Limits = new PdfReadLimits { MaxIndirectObjects = 5 } };

        PdfSanitizationResult result = PdfDocument.Load(source, readOptions).Sanitize();

        Assert.True(result.PreservationReport.IsPreserved, result.PreservationReport.Summary);
        PdfPageAction[] actions = result.ToDocument().Inspect().Pages[0].PageActions.ToArray();
        Assert.Equal(5, actions.Length);
        Assert.All(actions, static action => Assert.Equal("URI", action.ActionType));
    }

    [Fact]
    public void Sanitize_DoesNotRecountIndirectRetainedNextActionsDuringObjectSweep() {
        byte[] source = BuildForbiddenOpenActionWithRetainedNextPdf();
        var readOptions = new PdfLoadOptions { Limits = new PdfReadLimits { MaxWidgetActions = 1 } };

        PdfSanitizationResult result = PdfDocument.Load(source, readOptions).Sanitize();

        Assert.Equal("GoTo", result.ToDocument().Inspect().OpenAction?.ActionType);
        Assert.True(result.PreservationReport.IsPreserved, result.PreservationReport.Summary);
    }

    [Fact]
    public void Sanitize_FiltersSharedAllowedActionBeforeMarkingItsOriginalGraphNormalized() {
        PdfSanitizationResult result = PdfSanitizer.Sanitize(BuildSharedAllowedActionBeneathForbiddenRootPdf());

        Assert.Empty(result.RemainingFindings);
        Assert.Empty(PdfSanitizer.Analyze(result.ToBytes()));
        Assert.True(result.PreservationReport.IsPreserved, result.PreservationReport.Summary);
    }

    [Fact]
    public void Sanitize_CountsPromotedRetainedActionSiblingsOnce() {
        byte[] source = BuildForbiddenActionWithRetainedSiblingsPdf(actionCount: 5);
        var readOptions = new PdfLoadOptions { Limits = new PdfReadLimits { MaxIndirectObjects = 6 } };

        PdfSanitizationResult result = PdfDocument.Load(source, readOptions).Sanitize();

        PdfPageAction[] actions = result.ToDocument().Inspect().Pages[0].PageActions.ToArray();
        Assert.Equal(5, actions.Length);
        Assert.All(actions, static action => Assert.Equal("URI", action.ActionType));
        Assert.Contains(result.RemovedFindings, static finding => finding.Detail == "JavaScript");
    }

    [Fact]
    public void Sanitize_ReusesCustomReadLimitsForItsRewrittenArtifact() {
        byte[] source = BuildActiveContentPdf();
        var readOptions = new PdfLoadOptions {
            Limits = new PdfReadLimits {
                MaxInputBytes = source.LongLength,
                MaxJavaScripts = PdfReadLimits.DefaultMaxJavaScripts + 17
            }
        };

        PdfSanitizationResult result = PdfDocument.Load(source, readOptions).Sanitize();
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

    private static byte[] BuildForbiddenWidgetRootWithRetainedNextPdf() {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.7",
            "1 0 obj", "<< /Type /Catalog /Pages 2 0 R /AcroForm 5 0 R >>", "endobj",
            "2 0 obj", "<< /Type /Pages /Count 1 /Kids [3 0 R] >>", "endobj",
            "3 0 obj", "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 300] /Annots [6 0 R] >>", "endobj",
            "5 0 obj", "<< /Fields [6 0 R] >>", "endobj",
            "6 0 obj", "<< /Type /Annot /Subtype /Widget /FT /Btn /Ff 65536 /T (run) /Rect [20 20 120 44] /P 3 0 R /A 7 0 R >>", "endobj",
            "7 0 obj", "<< /S /JavaScript /JS (app.alert\\('remove'\\);) /Next 8 0 R >>", "endobj",
            "8 0 obj", "<< /S /GoTo /D [3 0 R /Fit] >>", "endobj",
            "trailer", "<< /Root 1 0 R /Size 9 >>", "%%EOF"
        });
        return Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildUnsafeWidgetUriPdf() {
        return Encoding.ASCII.GetBytes(string.Join("\n", new[] {
            "%PDF-1.7",
            "1 0 obj", "<< /Type /Catalog /Pages 2 0 R /AcroForm 5 0 R >>", "endobj",
            "2 0 obj", "<< /Type /Pages /Count 1 /Kids [3 0 R] >>", "endobj",
            "3 0 obj", "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 300] /Annots [6 0 R] >>", "endobj",
            "5 0 obj", "<< /Fields [6 0 R] >>", "endobj",
            "6 0 obj", "<< /Type /Annot /Subtype /Widget /FT /Btn /Ff 65536 /T (run) /Rect [20 20 120 44] /P 3 0 R /A << /S /URI /URI (javascript:unsafe) >> >>", "endobj",
            "trailer", "<< /Root 1 0 R /Size 7 >>", "%%EOF"
        }));
    }

    private static byte[] BuildAllowedRootWithForbiddenNextDescendantPdf() {
        string root = "<< /S /URI /URI (https://example.com/root) /Next [<< /S /JavaScript /JS (x) /Next << /S /URI /URI (https://example.com/promoted) >> >>] >>";
        return Encoding.ASCII.GetBytes(string.Join("\n", new[] {
            "%PDF-1.7",
            "1 0 obj", "<< /Type /Catalog /Pages 2 0 R >>", "endobj",
            "2 0 obj", "<< /Type /Pages /Count 1 /Kids [3 0 R] >>", "endobj",
            "3 0 obj", "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 4 0 R /AA << /O " + root + " >> >>", "endobj",
            "4 0 obj", "<< /Length 0 >>", "stream", string.Empty, "endstream", "endobj",
            "trailer", "<< /Root 1 0 R /Size 5 >>", "%%EOF"
        }));
    }

    private static byte[] BuildForbiddenViewerRootWithRetainedNextPdf(bool catalogAction) {
        string catalogEntry = catalogAction ? " /AA << /WC 7 0 R >>" : string.Empty;
        string pageEntry = catalogAction ? string.Empty : " /AA << /O 7 0 R >>";
        return Encoding.ASCII.GetBytes(string.Join("\n", new[] {
            "%PDF-1.7",
            "1 0 obj", "<< /Type /Catalog /Pages 2 0 R" + catalogEntry + " >>", "endobj",
            "2 0 obj", "<< /Type /Pages /Count 1 /Kids [3 0 R] >>", "endobj",
            "3 0 obj", "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 4 0 R" + pageEntry + " >>", "endobj",
            "4 0 obj", "<< /Length 0 >>", "stream", string.Empty, "endstream", "endobj",
            "7 0 obj", "<< /S /JavaScript /JS (app.alert\\('remove'\\);) /Next 8 0 R >>", "endobj",
            "8 0 obj", "<< /S /GoTo /D [3 0 R /Fit] >>", "endobj",
            "trailer", "<< /Root 1 0 R /Size 9 >>", "%%EOF"
        }));
    }

    private static byte[] BuildSharedRetainedActionDagPdf(int depth) {
        var lines = new List<string> {
            "%PDF-1.7",
            "1 0 obj", "<< /Type /Catalog /Pages 2 0 R >>", "endobj",
            "2 0 obj", "<< /Type /Pages /Count 1 /Kids [3 0 R] >>", "endobj",
            "3 0 obj", "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 4 0 R /AA << /O 7 0 R >> >>", "endobj",
            "4 0 obj", "<< /Length 0 >>", "stream", string.Empty, "endstream", "endobj"
        };
        for (int index = 0; index < depth; index++) {
            int objectNumber = 7 + index;
            string next = index + 1 < depth
                ? " /Next [" + (objectNumber + 1).ToString(System.Globalization.CultureInfo.InvariantCulture) + " 0 R " +
                    (objectNumber + 1).ToString(System.Globalization.CultureInfo.InvariantCulture) + " 0 R]"
                : string.Empty;
            string action = index == 0
                ? "<< /S /JavaScript /JS (remove)" + next + " >>"
                : "<< /S /URI /URI (https://example.test/)" + next + " >>";
            lines.Add(objectNumber.ToString(System.Globalization.CultureInfo.InvariantCulture) + " 0 obj");
            lines.Add(action);
            lines.Add("endobj");
        }
        lines.Add("trailer");
        lines.Add("<< /Root 1 0 R /Size " + (7 + depth).ToString(System.Globalization.CultureInfo.InvariantCulture) + " >>");
        lines.Add("%%EOF");
        return Encoding.ASCII.GetBytes(string.Join("\n", lines));
    }

    private static byte[] BuildLinearRetainedActionChainPdf(int actionCount) {
        string action = "<< /S /URI /URI (https://example.test/" + actionCount.ToString(System.Globalization.CultureInfo.InvariantCulture) + ") >>";
        for (int index = actionCount - 1; index > 0; index--) {
            action = "<< /S /URI /URI (https://example.test/" + index.ToString(System.Globalization.CultureInfo.InvariantCulture) + ") /Next [" + action + "] >>";
        }
        return Encoding.ASCII.GetBytes(string.Join("\n", new[] {
            "%PDF-1.7",
            "1 0 obj", "<< /Type /Catalog /Pages 2 0 R >>", "endobj",
            "2 0 obj", "<< /Type /Pages /Count 1 /Kids [3 0 R] >>", "endobj",
            "3 0 obj", "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 4 0 R /AA << /O " + action + " >> >>", "endobj",
            "4 0 obj", "<< /Length 0 >>", "stream", string.Empty, "endstream", "endobj",
            "trailer", "<< /Root 1 0 R /Size 5 >>", "%%EOF"
        }));
    }

    private static byte[] BuildForbiddenActionWithRetainedSiblingsPdf(int actionCount) {
        string siblings = string.Join(" ", Enumerable.Range(1, actionCount).Select(static index =>
            "<< /S /URI /URI (https://example.test/" + index.ToString(System.Globalization.CultureInfo.InvariantCulture) + ") >>"));
        return Encoding.ASCII.GetBytes(string.Join("\n", new[] {
            "%PDF-1.7",
            "1 0 obj", "<< /Type /Catalog /Pages 2 0 R >>", "endobj",
            "2 0 obj", "<< /Type /Pages /Count 1 /Kids [3 0 R] >>", "endobj",
            "3 0 obj", "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 4 0 R /AA << /O << /S /JavaScript /JS (remove) /Next [" + siblings + "] >> >> >>", "endobj",
            "4 0 obj", "<< /Length 0 >>", "stream", string.Empty, "endstream", "endobj",
            "trailer", "<< /Root 1 0 R /Size 5 >>", "%%EOF"
        }));
    }

    private static byte[] BuildForbiddenOpenActionWithRetainedNextPdf() {
        return Encoding.ASCII.GetBytes(string.Join("\n", new[] {
            "%PDF-1.7",
            "1 0 obj", "<< /Type /Catalog /Pages 2 0 R /OpenAction 7 0 R >>", "endobj",
            "2 0 obj", "<< /Type /Pages /Count 1 /Kids [3 0 R] >>", "endobj",
            "3 0 obj", "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 4 0 R >>", "endobj",
            "4 0 obj", "<< /Length 0 >>", "stream", string.Empty, "endstream", "endobj",
            "7 0 obj", "<< /S /JavaScript /JS (remove) /Next 8 0 R >>", "endobj",
            "8 0 obj", "<< /S /GoTo /D [3 0 R /Fit] >>", "endobj",
            "trailer", "<< /Root 1 0 R /Size 9 >>", "%%EOF"
        }));
    }

    private static byte[] BuildSharedAllowedActionBeneathForbiddenRootPdf() {
        return Encoding.ASCII.GetBytes(string.Join("\n", new[] {
            "%PDF-1.7",
            "1 0 obj", "<< /Type /Catalog /Pages 2 0 R /OpenAction 8 0 R >>", "endobj",
            "2 0 obj", "<< /Type /Pages /Count 1 /Kids [3 0 R] >>", "endobj",
            "3 0 obj", "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 4 0 R /AA << /O 7 0 R >> >>", "endobj",
            "4 0 obj", "<< /Length 0 >>", "stream", string.Empty, "endstream", "endobj",
            "7 0 obj", "<< /S /JavaScript /JS (remove-root) /Next 8 0 R >>", "endobj",
            "8 0 obj", "<< /S /URI /URI (https://example.test/keep) /Next 9 0 R >>", "endobj",
            "9 0 obj", "<< /S /JavaScript /JS (remove-child) >>", "endobj",
            "trailer", "<< /Root 1 0 R /Size 10 >>", "%%EOF"
        }));
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
            "1 0 obj\n<< /Type /Catalog /Pages 2 0 R /OpenAction [3 0 R /Fit] /AA << /WC << /S /URI /URI (https://example.com/catalog) >> /WS << /S /GoTo /D [3 0 R /Fit] >> /WP << /S /GoToR /F (remote.pdf) /D [0 /Fit] >> >> >>\nendobj\n" +
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
