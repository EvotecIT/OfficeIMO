using System.Text;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfSanitizerTests {
    [Fact]
    public void InspectAndSanitizeBeforeSharingUsesOneTypedPolicyAndVerifiedRewrite() {
        byte[] source = BuildBeforeSharingPdf();
        var policy = new PdfSanitizationOptions {
            ContentKindsToRemove = PdfSanitizationContentKind.All,
            ActionKindsToRemove = PdfSanitizationActionKind.All
        };

        PdfSanitizationReport preview = PdfDocument.Load(source).InspectSanitization(policy);

        Assert.Equal(6, preview.CategoryCounts.UserMetadata);
        Assert.Equal(1, preview.CategoryCounts.EmbeddedFiles);
        Assert.Equal(2, preview.CategoryCounts.Actions);
        Assert.Equal(1, preview.CategoryCounts.CommentsAndMarkup);
        Assert.Equal(1, preview.CategoryCounts.Bookmarks);
        Assert.Equal(1, preview.CategoryCounts.OptionalContent);
        Assert.Equal(12, preview.TotalCount);

        PdfSanitizationResult result = PdfDocument.Load(source).Sanitize(policy);
        PdfDocumentInfo info = result.ToDocument().Inspect();
        string raw = PdfEncoding.Latin1GetString(result.ToBytes());

        Assert.True(result.IsSanitized);
        Assert.Equal(12, result.RemovedCategoryCounts.Total);
        Assert.Equal(0, result.RemainingCategoryCounts.Total);
        Assert.Null(info.Metadata.Title);
        Assert.Null(info.Metadata.Author);
        Assert.Null(info.Metadata.Subject);
        Assert.Null(info.Metadata.Keywords);
        Assert.False(info.HasXmpMetadata);
        Assert.Empty(info.Attachments);
        Assert.Empty(info.Outlines);
        Assert.False(info.HasOptionalContent);
        Assert.DoesNotContain(info.Annotations, static annotation => annotation.Subtype == "Text" || annotation.Subtype == "FileAttachment");
        Assert.Contains(info.Annotations, static annotation => annotation.Subtype == "Link");
        Assert.Contains(info.Annotations, static annotation => annotation.Subtype == "Widget");
        Assert.Equal("OfficeIMO-Test", ReadInfoString(result.ToBytes(), "Producer"));
        Assert.Equal(new DateTimeOffset(2026, 1, 2, 3, 4, 5, TimeSpan.Zero), info.Metadata.CreationDate);
        Assert.Equal(new DateTimeOffset(2026, 2, 3, 4, 5, 6, TimeSpan.Zero), info.Metadata.ModificationDate);
        Assert.Equal(PdfTrappingStatus.False, info.Metadata.TrappingStatus);
        Assert.DoesNotContain("PRIVATE-", raw, StringComparison.Ordinal);
        Assert.Contains("VISIBLE-PAGE-CONTENT", result.ToDocument().Read().Text, StringComparison.Ordinal);
    }

    [Fact]
    public void SanitizeBeforeSharingCanRemoveOnlyUserMetadata() {
        byte[] source = BuildBeforeSharingPdf();
        var policy = new PdfSanitizationOptions {
            ContentKindsToRemove = PdfSanitizationContentKind.UserMetadata,
            ActionKindsToRemove = PdfSanitizationActionKind.All
        };

        PdfSanitizationResult result = PdfDocument.Load(source).Sanitize(policy);
        PdfDocumentInfo info = result.ToDocument().Inspect();

        Assert.True(result.IsSanitized);
        Assert.Equal(6, result.RemovedCategoryCounts.UserMetadata);
        Assert.Equal(0, result.RemovedCategoryCounts.Actions);
        Assert.Single(info.Attachments);
        Assert.Single(info.Outlines);
        Assert.True(info.HasOptionalContent);
        Assert.Contains(info.Annotations, static annotation => annotation.Subtype == "Text");
        Assert.Contains(PdfSanitizer.Analyze(result.ToBytes()), static finding => finding.ActionKind == PdfSanitizationActionKind.JavaScript);
    }

    [Theory]
    [InlineData(PdfSanitizationContentKind.UserMetadata)]
    [InlineData(PdfSanitizationContentKind.EmbeddedFiles)]
    [InlineData(PdfSanitizationContentKind.Actions)]
    [InlineData(PdfSanitizationContentKind.CommentsAndMarkup)]
    [InlineData(PdfSanitizationContentKind.Bookmarks)]
    [InlineData(PdfSanitizationContentKind.OptionalContent)]
    public void SanitizeBeforeSharingUsesAnExactContentCategorySelection(PdfSanitizationContentKind selected) {
        byte[] source = BuildBeforeSharingPdf();
        var allCategories = new PdfSanitizationOptions {
            ContentKindsToRemove = PdfSanitizationContentKind.All,
            ActionKindsToRemove = PdfSanitizationActionKind.All
        };
        var selectedCategory = new PdfSanitizationOptions {
            ContentKindsToRemove = selected,
            ActionKindsToRemove = PdfSanitizationActionKind.All
        };

        PdfSanitizationReport before = PdfDocument.Load(source).InspectSanitization(allCategories);
        PdfSanitizationResult result = PdfDocument.Load(source).Sanitize(selectedCategory);
        PdfSanitizationReport after = result.ToDocument().InspectSanitization(allCategories);

        Assert.True(result.IsSanitized);
        Assert.Equal(before.CategoryCounts.GetCount(selected), result.RemovedCategoryCounts.GetCount(selected));
        Assert.Equal(0, after.CategoryCounts.GetCount(selected));
        foreach (PdfSanitizationContentKind unselected in new[] {
                     PdfSanitizationContentKind.UserMetadata,
                     PdfSanitizationContentKind.EmbeddedFiles,
                     PdfSanitizationContentKind.Actions,
                     PdfSanitizationContentKind.CommentsAndMarkup,
                     PdfSanitizationContentKind.Bookmarks,
                     PdfSanitizationContentKind.OptionalContent
                 }.Where(kind => kind != selected)) {
            Assert.Equal(before.CategoryCounts.GetCount(unselected), after.CategoryCounts.GetCount(unselected));
        }
    }

    [Fact]
    public void SanitizationOptionsRejectUnsupportedContentKindBits() {
        Assert.Throws<ArgumentOutOfRangeException>(() => new PdfSanitizationOptions {
            ContentKindsToRemove = (PdfSanitizationContentKind)(1 << 20)
        });
    }

    [Fact]
    public void InspectSanitization_ReportsTypedActionCountsForTheDefaultPolicy() {
        PdfSanitizationReport report = PdfDocument.Load(BuildActiveContentPdf()).InspectSanitization();

        Assert.Equal(2, report.ActionCounts.JavaScript);
        Assert.Equal(1, report.ActionCounts.Uri);
        Assert.Equal(1, report.ActionCounts.Launch);
        Assert.Equal(1, report.ActionCounts.SubmitForm);
        Assert.Equal(1, report.ActionCounts.GoToR);
        Assert.Equal(1, report.ActionCounts.GoToE);
        Assert.Equal(1, report.ActionCounts.ImportData);
        Assert.Equal(8, report.ActionCounts.Total);
        Assert.Equal(report.Findings.Count, report.TotalCount);
    }

    [Theory]
    [InlineData(PdfSanitizationActionKind.JavaScript, 2)]
    [InlineData(PdfSanitizationActionKind.Launch, 1)]
    [InlineData(PdfSanitizationActionKind.SubmitForm, 1)]
    [InlineData(PdfSanitizationActionKind.GoToR, 1)]
    [InlineData(PdfSanitizationActionKind.GoToE, 1)]
    [InlineData(PdfSanitizationActionKind.ImportData, 1)]
    public void InspectSanitization_UsesAnExactActionKindSelection(PdfSanitizationActionKind selected, int expectedCount) {
        var policy = new PdfSanitizationOptions { ActionKindsToRemove = selected };

        PdfSanitizationReport report = PdfDocument.Load(BuildActiveContentPdf()).InspectSanitization(policy);

        Assert.Equal(expectedCount, report.ActionCounts.GetCount(selected));
        Assert.Equal(expectedCount, report.ActionCounts.Total);
        Assert.All(report.Findings.Where(static finding => finding.ActionKind.HasValue),
            finding => Assert.Equal(selected, finding.ActionKind));
    }

    [Fact]
    public void Sanitize_CanRemoveJavaScriptWithoutRemovingOtherActionKinds() {
        var policy = new PdfSanitizationOptions {
            ActionKindsToRemove = PdfSanitizationActionKind.JavaScript
        };

        PdfSanitizationResult result = PdfSanitizer.Sanitize(BuildActiveContentPdf(), policy);
        IReadOnlyList<PdfSanitizationFinding> preservedActions = PdfSanitizer.Analyze(result.ToBytes());

        Assert.True(result.IsSanitized);
        Assert.Equal(2, result.RemovedActionCounts.JavaScript);
        Assert.DoesNotContain("JavaScript", PdfEncoding.Latin1GetString(result.ToBytes()), StringComparison.Ordinal);
        Assert.Contains(preservedActions, static finding => finding.ActionKind == PdfSanitizationActionKind.Launch);
        Assert.Contains(preservedActions, static finding => finding.ActionKind == PdfSanitizationActionKind.SubmitForm);
        Assert.Contains(preservedActions, static finding => finding.ActionKind == PdfSanitizationActionKind.GoToR);
    }

    [Fact]
    public void Sanitize_CanRemoveEveryUriActionWithoutRemovingOtherActionKinds() {
        var policy = new PdfSanitizationOptions {
            ActionKindsToRemove = PdfSanitizationActionKind.Uri
        };

        PdfSanitizationResult result = PdfSanitizer.Sanitize(BuildActiveContentPdf(), policy);
        PdfDocumentInfo info = PdfInspector.Inspect(result.ToBytes());
        IReadOnlyList<PdfSanitizationFinding> preservedActions = PdfSanitizer.Analyze(result.ToBytes());

        Assert.True(result.IsSanitized);
        Assert.Equal(3, result.RemovedActionCounts.Uri);
        Assert.Empty(info.LinkAnnotations.Where(static link => link.Uri != null));
        Assert.DoesNotContain("base.example", PdfEncoding.Latin1GetString(result.ToBytes()), StringComparison.Ordinal);
        Assert.Contains(preservedActions, static finding => finding.ActionKind == PdfSanitizationActionKind.JavaScript);
        Assert.Contains(preservedActions, static finding => finding.ActionKind == PdfSanitizationActionKind.Launch);
    }

    [Fact]
    public void Sanitize_ExplicitAllowListOverridesAnExactActionKindSelection() {
        var policy = new PdfSanitizationOptions {
            ActionKindsToRemove = PdfSanitizationActionKind.JavaScript |
                                  PdfSanitizationActionKind.Launch
        };
        policy.AllowedActionTypes.Add("JavaScript");

        PdfSanitizationReport preview = PdfDocument.Load(BuildActiveContentPdf()).InspectSanitization(policy);
        PdfSanitizationResult result = PdfSanitizer.Sanitize(BuildActiveContentPdf(), policy);
        IReadOnlyList<PdfSanitizationFinding> defaultPolicyFindings = PdfSanitizer.Analyze(result.ToBytes());

        Assert.Equal(0, preview.ActionCounts.JavaScript);
        Assert.Equal(1, preview.ActionCounts.Launch);
        Assert.Equal(0, result.RemovedActionCounts.JavaScript);
        Assert.Equal(1, result.RemovedActionCounts.Launch);
        Assert.Contains(defaultPolicyFindings, static finding => finding.ActionKind == PdfSanitizationActionKind.JavaScript);
        Assert.DoesNotContain(defaultPolicyFindings, static finding => finding.ActionKind == PdfSanitizationActionKind.Launch);
    }

    [Fact]
    public void Sanitize_DefaultUriSchemePolicyOverridesTheActionTypeAllowList() {
        var policy = new PdfSanitizationOptions();
        policy.AllowedActionTypes.Add("URI");

        PdfSanitizationReport preview = PdfDocument.Load(BuildUnsafeWidgetUriPdf()).InspectSanitization(policy);
        PdfSanitizationResult result = PdfSanitizer.Sanitize(BuildUnsafeWidgetUriPdf(), policy);

        Assert.Equal(1, preview.ActionCounts.Uri);
        Assert.Equal(1, result.RemovedActionCounts.Uri);
        Assert.DoesNotContain("javascript:unsafe", PdfEncoding.Latin1GetString(result.ToBytes()), StringComparison.Ordinal);
    }

    [Fact]
    public void InspectSanitization_InventoriesMalformedNamedJavaScriptContainer() {
        byte[] source = BuildMalformedNamedJavaScriptPdf();
        var policy = new PdfSanitizationOptions {
            ActionKindsToRemove = PdfSanitizationActionKind.JavaScript
        };

        PdfSanitizationReport preview = PdfDocument.Load(source).InspectSanitization(policy);
        PdfSanitizationResult result = PdfSanitizer.Sanitize(source, policy);

        Assert.Equal(1, preview.ActionCounts.JavaScript);
        Assert.Equal(1, result.RemovedActionCounts.JavaScript);
        Assert.True(result.IsSanitized);
        Assert.DoesNotContain("/JavaScript", PdfEncoding.Latin1GetString(result.ToBytes()), StringComparison.Ordinal);
    }

    [Fact]
    public void InspectSanitization_BoundsNamedJavaScriptReferenceChains() {
        byte[] source = BuildNamedJavaScriptReferenceChainPdf(referenceCount: 8);
        var readOptions = new PdfLoadOptions {
            Limits = new PdfReadLimits { MaxNameTreeDepth = 4 }
        };

        PdfReadLimitException exception = Assert.Throws<PdfReadLimitException>(() =>
            PdfDocument.Load(source, readOptions).InspectSanitization());

        Assert.Equal(PdfReadLimitKind.NameTreeDepth, exception.Kind);
    }

    [Fact]
    public void InspectSanitization_CountsDanglingNamedJavaScriptReferencesAgainstNodeLimit() {
        byte[] source = Encoding.ASCII.GetBytes(string.Join("\n", new[] {
            "%PDF-1.7",
            "1 0 obj", "<< /Type /Catalog /Pages 2 0 R /Names << /JavaScript << /Kids [100 0 R 101 0 R 102 0 R] >> >> >>", "endobj",
            "2 0 obj", "<< /Type /Pages /Count 1 /Kids [3 0 R] >>", "endobj",
            "3 0 obj", "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 4 0 R >>", "endobj",
            "4 0 obj", "<< /Length 0 >>", "stream", string.Empty, "endstream", "endobj",
            "trailer", "<< /Root 1 0 R /Size 5 >>", "%%EOF"
        }));
        var readOptions = new PdfLoadOptions {
            Limits = new PdfReadLimits { MaxNameTreeNodes = 2 }
        };

        PdfReadLimitException exception = Assert.Throws<PdfReadLimitException>(() =>
            PdfDocument.Load(source, readOptions).InspectSanitization());

        Assert.Equal(PdfReadLimitKind.NameTreeNodes, exception.Kind);
    }

    [Fact]
    public void InspectSanitization_CachesSharedNamedJavaScriptRootsAcrossTheScan() {
        byte[] source = BuildSharedNamedJavaScriptRootPdf();
        var readOptions = new PdfLoadOptions {
            Limits = new PdfReadLimits { MaxNameTreeNodes = 2 }
        };

        PdfSanitizationReport report = PdfDocument.Load(source, readOptions).InspectSanitization();

        Assert.Equal(2, report.ActionCounts.JavaScript);
    }

    [Fact]
    public void InspectSanitization_DoesNotChargeActionPayloadsToTheNameTreeNodeLimit() {
        byte[] source = Encoding.ASCII.GetBytes(string.Join("\n", new[] {
            "%PDF-1.7",
            "1 0 obj", "<< /Type /Catalog /Pages 2 0 R /Names << /JavaScript 5 0 R >> >>", "endobj",
            "2 0 obj", "<< /Type /Pages /Count 1 /Kids [3 0 R] >>", "endobj",
            "3 0 obj", "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 4 0 R >>", "endobj",
            "4 0 obj", "<< /Length 0 >>", "stream", string.Empty, "endstream", "endobj",
            "5 0 obj", "<< /Names [(Run) 6 0 R] >>", "endobj",
            "6 0 obj", "<< /S /JavaScript /JS 7 0 R >>", "endobj",
            "7 0 obj", "(app.alert\\('test'\\);)", "endobj",
            "trailer", "<< /Root 1 0 R /Size 8 >>", "%%EOF"
        }));
        var readOptions = new PdfLoadOptions {
            Limits = new PdfReadLimits { MaxNameTreeNodes = 1 }
        };

        PdfSanitizationReport report = PdfDocument.Load(source, readOptions).InspectSanitization();

        Assert.Equal(1, report.ActionCounts.JavaScript);
    }

    [Fact]
    public void SanitizationOptions_RejectUnsupportedActionKindBits() {
        Assert.Throws<ArgumentOutOfRangeException>(() => new PdfSanitizationOptions {
            ActionKindsToRemove = (PdfSanitizationActionKind)(1 << 20)
        });
    }

    [Fact]
    public void Analyze_HonorsPolicyCancellation() {
        byte[] source = BuildActiveContentPdf();
        using var cancellation = new System.Threading.CancellationTokenSource();
        cancellation.Cancel();

        Assert.Throws<OperationCanceledException>(() => PdfSanitizer.Analyze(
            source,
            new PdfSanitizationOptions { CancellationToken = cancellation.Token }));
    }

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
            "<< /Type /Catalog /Pages 2 0 R /Names << /JavaScript << /Names [(Open) 6 0 R] >> >> /URI << /Base (https://base.example/) >> /AA << /WC 12 0 R /WS 13 0 R /WP 14 0 R >> >>",
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

    private static byte[] BuildBeforeSharingPdf() {
        const string pageContent = "BT /F1 12 Tf 20 150 Td (VISIBLE-PAGE-CONTENT) Tj ET\n/OC /Layer BDC BT /F1 12 Tf 20 120 Td (VISIBLE-AFTER-LAYER-FLATTEN) Tj ET EMC";
        const string xmp = "<?xpacket begin=''?><x:xmpmeta xmlns:x='adobe:ns:meta/'><rdf:RDF xmlns:rdf='http://www.w3.org/1999/02/22-rdf-syntax-ns#'><rdf:Description rdf:about='' xmlns:dc='http://purl.org/dc/elements/1.1/'><dc:title><rdf:Alt><rdf:li xml:lang='x-default'>PRIVATE-XMP-TITLE</rdf:li></rdf:Alt></dc:title></rdf:Description></rdf:RDF></x:xmpmeta><?xpacket end='w'?>";
        const string payload = "PRIVATE-ATTACHMENT-PAYLOAD";
        string pdf = string.Join("\n", new[] {
            "%PDF-1.7",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R /Outlines 10 0 R /PageMode /UseOutlines /Names << /EmbeddedFiles << /Names [(payload.txt) 14 0 R] >> /JavaScript << /Names [(startup) 17 0 R] >> >> /Metadata 12 0 R /OCProperties << /OCGs [16 0 R] /D << /ON [16 0 R] /Order [16 0 R] >> >> >>",
            "endobj",
            "2 0 obj", "<< /Type /Pages /Count 1 /Kids [3 0 R] >>", "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 240 180] /Resources << /Font << /F1 8 0 R >> /Properties << /Layer 16 0 R >> >> /Contents 4 0 R /Annots [5 0 R 6 0 R 7 0 R 15 0 R] >>",
            "endobj",
            StreamObject(4, string.Empty, pageContent),
            "5 0 obj", "<< /Type /Annot /Subtype /Text /Rect [20 20 40 40] /Contents (PRIVATE-COMMENT) >>", "endobj",
            "6 0 obj", "<< /Type /Annot /Subtype /Link /Rect [50 20 100 40] /A << /S /URI /URI (https://example.com/) >> >>", "endobj",
            "7 0 obj", "<< /Type /Annot /Subtype /Widget /Rect [110 20 180 40] >>", "endobj",
            "8 0 obj", "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>", "endobj",
            "10 0 obj", "<< /Type /Outlines /First 11 0 R /Last 11 0 R /Count 1 >>", "endobj",
            "11 0 obj", "<< /Title (PRIVATE-BOOKMARK) /Parent 10 0 R /Dest [3 0 R /Fit] >>", "endobj",
            StreamObject(12, "/Type /Metadata /Subtype /XML", xmp),
            StreamObject(13, "/Type /EmbeddedFile /Subtype /text#2Fplain", payload),
            "14 0 obj", "<< /Type /Filespec /F (payload.txt) /UF (payload.txt) /EF << /F 13 0 R /UF 13 0 R >> >>", "endobj",
            "15 0 obj", "<< /Type /Annot /Subtype /FileAttachment /Rect [190 20 210 40] /FS 14 0 R /Contents (PRIVATE-ATTACHMENT-COMMENT) >>", "endobj",
            "16 0 obj", "<< /Type /OCG /Name (PRIVATE-LAYER) >>", "endobj",
            "17 0 obj", "<< /S /JavaScript /JS (app.alert('PRIVATE-SCRIPT')) >>", "endobj",
            "20 0 obj",
            "<< /Title (PRIVATE-TITLE) /Author (PRIVATE-AUTHOR) /Subject (PRIVATE-SUBJECT) /Keywords (PRIVATE-KEYWORDS) /Creator (PRIVATE-CREATOR) /Producer (OfficeIMO-Test) /CreationDate (D:20260102030405Z) /ModDate (D:20260203040506Z) /Trapped /False >>",
            "endobj",
            "trailer", "<< /Root 1 0 R /Info 20 0 R /Size 21 >>", "%%EOF"
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

    private static byte[] BuildMalformedNamedJavaScriptPdf() {
        return Encoding.ASCII.GetBytes(string.Join("\n", new[] {
            "%PDF-1.7",
            "1 0 obj", "<< /Type /Catalog /Pages 2 0 R /Names << /JavaScript << /Names [] >> >> >>", "endobj",
            "2 0 obj", "<< /Type /Pages /Count 1 /Kids [3 0 R] >>", "endobj",
            "3 0 obj", "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 4 0 R >>", "endobj",
            "4 0 obj", "<< /Length 0 >>", "stream", string.Empty, "endstream", "endobj",
            "trailer", "<< /Root 1 0 R /Size 5 >>", "%%EOF"
        }));
    }

    private static byte[] BuildNamedJavaScriptReferenceChainPdf(int referenceCount) {
        var lines = new List<string> {
            "%PDF-1.7",
            "1 0 obj", "<< /Type /Catalog /Pages 2 0 R /Names << /JavaScript 5 0 R >> >>", "endobj",
            "2 0 obj", "<< /Type /Pages /Count 1 /Kids [3 0 R] >>", "endobj",
            "3 0 obj", "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 4 0 R >>", "endobj",
            "4 0 obj", "<< /Length 0 >>", "stream", string.Empty, "endstream", "endobj"
        };
        for (int index = 0; index < referenceCount; index++) {
            int objectNumber = 5 + index;
            lines.Add(objectNumber.ToString(System.Globalization.CultureInfo.InvariantCulture) + " 0 obj");
            lines.Add(index + 1 < referenceCount
                ? "<< /Kids [" + (objectNumber + 1).ToString(System.Globalization.CultureInfo.InvariantCulture) + " 0 R] >>"
                : "<< /Names [] >>");
            lines.Add("endobj");
        }
        lines.Add("trailer");
        lines.Add("<< /Root 1 0 R /Size " + (5 + referenceCount).ToString(System.Globalization.CultureInfo.InvariantCulture) + " >>");
        lines.Add("%%EOF");
        return Encoding.ASCII.GetBytes(string.Join("\n", lines));
    }

    private static byte[] BuildSharedNamedJavaScriptRootPdf() {
        return Encoding.ASCII.GetBytes(string.Join("\n", new[] {
            "%PDF-1.7",
            "1 0 obj", "<< /Type /Catalog /Pages 2 0 R /Names << /JavaScript 5 0 R /Also << /JavaScript 5 0 R >> >> >>", "endobj",
            "2 0 obj", "<< /Type /Pages /Count 1 /Kids [3 0 R] >>", "endobj",
            "3 0 obj", "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 4 0 R >>", "endobj",
            "4 0 obj", "<< /Length 0 >>", "stream", string.Empty, "endstream", "endobj",
            "5 0 obj", "<< /Kids [6 0 R] >>", "endobj",
            "6 0 obj", "<< /Names [] >>", "endobj",
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

    private static string StreamObject(int objectNumber, string dictionaryEntries, string content) =>
        objectNumber.ToString(System.Globalization.CultureInfo.InvariantCulture) + " 0 obj\n<< " + dictionaryEntries +
        " /Length " + Encoding.ASCII.GetByteCount(content).ToString(System.Globalization.CultureInfo.InvariantCulture) +
        " >>\nstream\n" + content + "\nendstream\nendobj";

    private static string? ReadInfoString(byte[] pdf, string key) {
        var (objects, trailerRaw) = PdfSyntax.ParseObjects(pdf);
        if (!PdfSyntax.TryGetTrailerReference(trailerRaw, "Info", limits: null, out PdfReference infoReference) ||
            !objects.TryGetValue(infoReference.ObjectNumber, out PdfIndirectObject? infoObject) ||
            infoObject.Value is not PdfDictionary info ||
            info.Get<PdfStringObj>(key) is not PdfStringObj value) {
            return null;
        }
        return value.Value;
    }
}
