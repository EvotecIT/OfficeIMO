using System;
using System.Collections.Generic;
using System.Linq;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfRewritePreservationMatrixTests {
    [Fact]
    public void Assess_ClassifiesSafeFailedBlockedAndOperationFailedRows() {
        byte[] preservationSource = PdfRewritePreservationTestSupport.BuildPreservationProofPdf();
        IReadOnlyList<PdfRewritePreservationMatrixScenario> premiumScenarios =
            PdfRewritePreservationMatrixScenarioSupport.BuildPremiumFeatureScenarios();

        var operationFailed = new PdfRewritePreservationMatrixScenario(
                "unexpected-operation-failure",
                "SyntheticFailure",
                preservationSource,
                _ => throw new InvalidOperationException("Synthetic matrix failure.")) {
                ExpectedClassification = PdfRewritePreservationMatrixClassification.OperationFailed
            }
            .WithSourceFeatures("diagnostics");

        PdfRewritePreservationMatrixReport report = PdfRewritePreservationMatrix.AssertExpected(
            premiumScenarios.Concat(new[] { operationFailed }));

        Assert.True(report.Passed);
        Assert.Equal(PdfRewritePreservationMatrixScenarioSupport.PremiumScenarioCount + 1, report.Entries.Count);

        PdfRewritePreservationMatrixEntry safeEntry = Assert.Single(report.Entries, entry => entry.Id == "metadata-update-safe");
        Assert.True(safeEntry.Passed);
        Assert.Equal(PdfRewritePreservationMatrixClassification.RewriteSafe, safeEntry.ActualClassification);
        Assert.NotNull(safeEntry.PreservationReport);
        Assert.True(safeEntry.PreservationReport!.IsPreserved);
        Assert.Contains("attachments", safeEntry.SourceFeatures);

        PdfRewritePreservationMatrixEntry driftEntry = Assert.Single(report.Entries, entry => entry.Id == "source-structure-drift-detected");
        Assert.Equal(PdfRewritePreservationMatrixClassification.PreservationFailed, driftEntry.ActualClassification);
        Assert.NotNull(driftEntry.PreservationReport);
        Assert.Contains(driftEntry.PreservationReport!.Issues, issue => issue.Feature == "SourceStructure.XrefStreams");
        Assert.Contains("SourceStructure.XrefStreams", driftEntry.Summary, StringComparison.Ordinal);

        PdfRewritePreservationMatrixEntry optionalContentEntry = Assert.Single(report.Entries, entry => entry.Id == "optional-content-drift-detected");
        Assert.Equal(PdfRewritePreservationMatrixClassification.PreservationFailed, optionalContentEntry.ActualClassification);
        Assert.NotNull(optionalContentEntry.PreservationReport);
        Assert.Contains(optionalContentEntry.PreservationReport!.Issues, issue => issue.Feature == "OptionalContent.Groups[0].Name");
        Assert.Contains("optional-content", optionalContentEntry.SourceFeatures);

        PdfRewritePreservationMatrixEntry formFillEntry = Assert.Single(report.Entries, entry => entry.Id == "form-fill-safe");
        Assert.Equal(PdfRewritePreservationMatrixClassification.RewriteSafe, formFillEntry.ActualClassification);
        Assert.NotNull(formFillEntry.PreservationReport);
        Assert.True(formFillEntry.PreservationReport!.Rewritten.HasForms);
        Assert.Contains("appearances", formFillEntry.SourceFeatures);

        PdfRewritePreservationMatrixEntry formsBlockedEntry = Assert.Single(report.Entries, entry => entry.Id == "forms-full-rewrite-blocked");
        Assert.Equal(PdfRewritePreservationMatrixClassification.Blocked, formsBlockedEntry.ActualClassification);
        Assert.Contains("PDF form fields are not supported for rewriting", formsBlockedEntry.FailureMessage, StringComparison.Ordinal);

        PdfRewritePreservationMatrixEntry taggedBlockedEntry = Assert.Single(report.Entries, entry => entry.Id == "tagged-full-rewrite-blocked");
        Assert.Equal(PdfRewritePreservationMatrixClassification.Blocked, taggedBlockedEntry.ActualClassification);
        Assert.Contains("PDF tagged content structure is not supported for rewriting", taggedBlockedEntry.FailureMessage, StringComparison.Ordinal);

        PdfRewritePreservationMatrixEntry activeBlockedEntry = Assert.Single(report.Entries, entry => entry.Id == "active-content-full-rewrite-blocked");
        Assert.Equal(PdfRewritePreservationMatrixClassification.Blocked, activeBlockedEntry.ActualClassification);
        Assert.Contains("PDF active content is not supported for rewriting", activeBlockedEntry.FailureMessage, StringComparison.Ordinal);

        PdfRewritePreservationMatrixEntry blockedEntry = Assert.Single(report.Entries, entry => entry.Id == "signed-full-rewrite-blocked");
        Assert.Equal(PdfRewritePreservationMatrixClassification.Blocked, blockedEntry.ActualClassification);
        Assert.Equal("NotSupportedException", blockedEntry.FailureType);
        Assert.Contains("Signed PDF files are not supported for rewriting", blockedEntry.FailureMessage, StringComparison.Ordinal);
        Assert.Contains("signature", blockedEntry.SourceFeatures);

        PdfRewritePreservationMatrixEntry failedEntry = Assert.Single(report.Entries, entry => entry.Id == "unexpected-operation-failure");
        Assert.Equal(PdfRewritePreservationMatrixClassification.OperationFailed, failedEntry.ActualClassification);
        Assert.Equal("InvalidOperationException", failedEntry.FailureType);
        Assert.Contains("Synthetic matrix failure", failedEntry.Summary, StringComparison.Ordinal);

        PdfRewritePreservationMatrixSummary summary = report.ToSummary();
        Assert.True(summary.Passed);
        Assert.Equal(report.Summary, summary.Summary);
        Assert.Equal(report.Entries.Count, summary.Rows.Count);
        PdfRewritePreservationMatrixRowSummary optionalSummary =
            Assert.Single(summary.Rows, row => row.Id == "optional-content-drift-detected");
        Assert.Equal("PreservationFailed", optionalSummary.ActualClassification);
        Assert.NotNull(optionalSummary.Issues);
        Assert.Contains(optionalSummary.Issues!, issue => issue.Feature == "OptionalContent.Groups[0].Name");
    }

    [Fact]
    public void AssertExpected_ThrowsWhenObservedClassificationDiffersFromExpectation() {
        byte[] source = PdfRewritePreservationTestSupport.BuildSignedIncrementalProofPdf();
        var scenario = new PdfRewritePreservationMatrixScenario(
            "signed-should-not-be-safe",
            "MetadataUpdate",
            source,
            pdf => PdfMetadataEditor.UpdateMetadata(pdf, title: "Unexpected"));

        var exception = Assert.Throws<InvalidOperationException>(() => PdfRewritePreservationMatrix.AssertExpected(new[] { scenario }));

        Assert.Contains("signed-should-not-be-safe", exception.Message, StringComparison.Ordinal);
        Assert.Contains("Blocked", exception.Message, StringComparison.Ordinal);
    }
}
