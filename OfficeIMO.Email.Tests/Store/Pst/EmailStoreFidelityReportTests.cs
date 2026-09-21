using OfficeIMO.Email.Store;
using Xunit;

namespace OfficeIMO.Email.Tests;

public sealed class EmailStoreFidelityReportTests {
    [Fact]
    public void StoreFormatExportReportsImplementCommonTypedFidelityContract() {
        var reference = new EmailStoreItemReference("item-1", "folder-1", false, false);
        var skipped = new EmailStoreDiagnostic(
            "EMAIL_STORE_TEST_EXPORT_SKIPPED",
            "The selected source item was not exported.",
            EmailStoreDiagnosticSeverity.Warning,
            "item/item-1",
            operation: "export",
            byteOffset: null,
            limitName: null,
            actualValue: null,
            maximumValue: null,
            disposition: EmailDiagnosticDisposition.Skipped,
            dataLossRisk: EmailDataLossRisk.Confirmed,
            suggestedAction: null);
        var failedEntry = new EmailStoreExportEntry(reference, null, 0, new[] { skipped });
        var directory = new EmailStoreExportReport(
            "export", true, manifestPath: null, new[] { failedEntry }, Array.Empty<EmailStoreDiagnostic>());
        var mbox = new EmailStoreMboxExportReport(
            "export.mbox", true, Array.Empty<EmailStoreMboxExportEntry>(), Array.Empty<EmailStoreDiagnostic>());
        var discovery = new EmailStoreRecoveryReport(
            itemsScanned: 1,
            stoppedAtLimit: true,
            recoveredItems: Array.Empty<EmailStoreItemReference>(),
            diagnostics: Array.Empty<EmailStoreDiagnostic>());
        var recovery = new EmailStoreRecoveryExportReport(
            "recovery", discovery, manifestPath: null,
            entries: Array.Empty<EmailStoreExportEntry>(),
            diagnostics: Array.Empty<EmailStoreDiagnostic>());

        IOfficeConversionReport[] reports = { directory, mbox, recovery };
        IReadOnlyList<OfficeConversionFidelityDiagnostic> flattened =
            OfficeConversionFidelityDiagnostics.Flatten(reports);

        Assert.All(reports, report => {
            Assert.True(report.HasLoss);
            Assert.Throws<InvalidDataException>(report.RequireNoLoss);
        });
        Assert.Contains(flattened, diagnostic =>
            diagnostic.Code == "EMAIL_STORE_TEST_EXPORT_SKIPPED" &&
            diagnostic.LossKind == OfficeConversionLossKind.Omission);
        Assert.Contains(flattened, diagnostic =>
            diagnostic.Code == "EMAIL_STORE_EXPORT_SELECTION_TRUNCATED" &&
            diagnostic.LossKind == OfficeConversionLossKind.Omission);
        Assert.Contains(flattened, diagnostic =>
            diagnostic.Code == "EMAIL_STORE_EXPORT_ITEMS_OMITTED" &&
            diagnostic.LossKind == OfficeConversionLossKind.Omission);
        Assert.Contains(flattened, diagnostic =>
            diagnostic.Code == "EMAIL_STORE_MBOX_SELECTION_TRUNCATED" &&
            diagnostic.LossKind == OfficeConversionLossKind.Omission);
        Assert.Contains(flattened, diagnostic =>
            diagnostic.Code == "EMAIL_STORE_RECOVERY_DISCOVERY_TRUNCATED" &&
            diagnostic.LossKind == OfficeConversionLossKind.Omission);
    }

    [Fact]
    public void SuccessfulStoreFormatExportReportsPassStrictAcceptance() {
        var directory = new EmailStoreExportReport(
            "export", false, manifestPath: null,
            entries: Array.Empty<EmailStoreExportEntry>(),
            diagnostics: Array.Empty<EmailStoreDiagnostic>());
        var mbox = new EmailStoreMboxExportReport(
            "export.mbox", false,
            entries: Array.Empty<EmailStoreMboxExportEntry>(),
            diagnostics: Array.Empty<EmailStoreDiagnostic>());
        var discovery = new EmailStoreRecoveryReport(
            itemsScanned: 0,
            stoppedAtLimit: false,
            recoveredItems: Array.Empty<EmailStoreItemReference>(),
            diagnostics: Array.Empty<EmailStoreDiagnostic>());
        var recovery = new EmailStoreRecoveryExportReport(
            "recovery", discovery, manifestPath: null,
            entries: Array.Empty<EmailStoreExportEntry>(),
            diagnostics: Array.Empty<EmailStoreDiagnostic>());

        foreach (IOfficeConversionReport report in new IOfficeConversionReport[] { directory, mbox, recovery }) {
            Assert.False(report.HasLoss);
            Assert.Empty(report.FidelityDiagnostics);
            report.RequireNoLoss();
        }
    }

    [Fact]
    public void StoreWriteReport_ClassifiesContinuedAttachmentLossAsOmission() {
        var omitted = new EmailStoreDiagnostic(
            "EMAIL_STORE_PST_WRITE_ATTACHMENT_CONTENT_UNAVAILABLE",
            "Attachment content was unavailable and only metadata was written.",
            EmailStoreDiagnosticSeverity.Error,
            "attachment/0x00000001",
            operation: "write",
            byteOffset: null,
            limitName: null,
            actualValue: null,
            maximumValue: null,
            disposition: EmailDiagnosticDisposition.Skipped,
            dataLossRisk: EmailDataLossRisk.Confirmed,
            suggestedAction: null);
        var report = new EmailStorePstWriteReport(
            "destination.pst", 1, 1, 128, new[] { omitted });

        OfficeConversionFidelityDiagnostic diagnostic = Assert.Single(report.FidelityDiagnostics);

        Assert.Equal(OfficeConversionLossKind.Omission, diagnostic.LossKind);
        Assert.True(report.HasLoss);
        Assert.Throws<InvalidDataException>(report.RequireNoLoss);
    }

    [Fact]
    public void StoreWriteReportPreservesOmissionAndFailureCategories() {
        var skipped = new EmailStoreDiagnostic(
            "EMAIL_STORE_TEST_SKIPPED",
            "One source item was skipped.",
            EmailStoreDiagnosticSeverity.Warning,
            "folder/item",
            "write",
            byteOffset: null,
            limitName: null,
            actualValue: null,
            maximumValue: null,
            EmailDiagnosticDisposition.Skipped,
            EmailDataLossRisk.Confirmed,
            suggestedAction: null);
        var failed = new EmailStoreDiagnostic(
            "EMAIL_STORE_TEST_FAILED",
            "The destination could not preserve the item.",
            EmailStoreDiagnosticSeverity.Error,
            "folder/item");
        var report = new EmailStorePstWriteReport(
            "destination.pst", 1, 0, 0, new[] { skipped, failed });

        IOfficeConversionReport common = report;

        Assert.True(report.HasDataLoss);
        Assert.True(common.HasLoss);
        Assert.Equal(OfficeConversionLossKind.Omission,
            Assert.Single(common.FidelityDiagnostics,
                diagnostic => diagnostic.Code == "EMAIL_STORE_TEST_SKIPPED").LossKind);
        Assert.Equal(OfficeConversionLossKind.Failure,
            Assert.Single(common.FidelityDiagnostics,
                diagnostic => diagnostic.Code == "EMAIL_STORE_TEST_FAILED").LossKind);
        Assert.Throws<InvalidDataException>(common.RequireNoLoss);
    }

    [Fact]
    public void StoreConversionReportClassifiesSkippedItemsWithoutAggregateOnlyLoss() {
        var write = new EmailStorePstWriteReport(
            "destination.pst", 1, 1, 128, Array.Empty<EmailStoreDiagnostic>());
        var identity = new EmailStoreSourceIdentity(
            EmailStoreFormat.Pst, 128, "catalog", "durable");
        var report = new EmailStorePstConversionReport(
            EmailStoreFormat.Pst,
            write,
            sourceFolders: 1,
            convertedItems: 1,
            skippedItems: 1,
            verification: null,
            diagnostics: Array.Empty<EmailStoreDiagnostic>(),
            sourceIdentity: identity,
            wasResumed: false);

        IOfficeConversionReport common = report;
        OfficeConversionFidelityDiagnostic diagnostic = Assert.Single(common.FidelityDiagnostics);

        Assert.Equal(EmailStoreMigrationDisposition.CompletedWithAcceptedLoss, report.Disposition);
        Assert.True(report.HasDataLoss);
        Assert.Equal("EMAIL_STORE_ITEMS_SKIPPED", diagnostic.Code);
        Assert.Equal(OfficeConversionLossKind.Omission, diagnostic.LossKind);
        Assert.Throws<InvalidDataException>(common.RequireNoLoss);
    }

    [Fact]
    public void StoreComposedReportsFailClosedOnVerificationMismatch() {
        var write = new EmailStorePstWriteReport(
            "destination.pst", 1, 1, 128, Array.Empty<EmailStoreDiagnostic>());
        var sourceIdentity = new EmailStoreSourceIdentity(
            EmailStoreFormat.Pst, 128, "catalog", "durable");
        var failedVerification = new EmailStorePstVerificationReport(
            attemptedItems: 1,
            matchedItems: 0,
            mismatchedItems: 1,
            failedItems: 0,
            issues: Array.Empty<EmailStorePstVerificationIssue>(),
            issuesTruncated: false,
            manifestPath: null);
        var conversion = new EmailStorePstConversionReport(
            EmailStoreFormat.Pst,
            write,
            sourceFolders: 1,
            convertedItems: 1,
            skippedItems: 0,
            verification: failedVerification,
            diagnostics: Array.Empty<EmailStoreDiagnostic>(),
            sourceIdentity,
            wasResumed: false);
        var compactionPlan = new EmailStorePstCompactionPlan(
            "compacted.pst",
            new EmailStorePstCompactionOptions(),
            sourceBytes: 128,
            itemsScanned: 2,
            selectedItems: 2,
            associatedItems: 0,
            orphanedItems: 0,
            excludedSearchFolderItems: 0,
            unknownSizeItems: 0,
            estimatedOutputBytes: 64,
            itemLimitReached: false,
            diagnostics: Array.Empty<EmailStoreDiagnostic>());
        var compaction = new EmailStorePstCompactionReport(compactionPlan, conversion);

        var mutationPlan = new EmailStorePstMutationPlan(
            "source.pst",
            Array.Empty<EmailStorePstMutationPlanOperation>(),
            resultingFolderCount: 1,
            resultingItemCount: 1,
            estimatedRewriteBytes: 128,
            diagnostics: Array.Empty<EmailStoreDiagnostic>());
        var mutationVerification = new EmailStorePstMutationVerificationReport(
            attemptedFolders: 1,
            matchedFolders: 0,
            mismatchedFolders: 1,
            failedFolders: 0,
            unexpectedFolders: 0,
            attemptedItems: 1,
            matchedItems: 0,
            mismatchedItems: 1,
            failedItems: 0,
            issues: Array.Empty<EmailStorePstMutationVerificationIssue>(),
            issuesTruncated: false);
        var mutation = new EmailStorePstMutationReport(
            "source.pst",
            backupPath: null,
            mutationPlan,
            write,
            mutationVerification,
            createdFolders: 0,
            renamedFolders: 0,
            movedFolders: 0,
            deletedFolders: 0,
            addedItems: 0,
            copiedItems: 0,
            replacedItems: 0,
            patchedItems: 0,
            movedItems: 0,
            deletedItems: 0,
            folderIdMap: new Dictionary<string, string>(),
            itemIdMap: new Dictionary<string, string>(),
            operationResults: Array.Empty<EmailStorePstMutationOperationResult>(),
            diagnostics: Array.Empty<EmailStoreDiagnostic>());

        foreach (IOfficeConversionReport report in new IOfficeConversionReport[] { conversion, compaction, mutation }) {
            Assert.True(report.HasLoss);
            Assert.Contains(report.FidelityDiagnostics,
                diagnostic => diagnostic.LossKind == OfficeConversionLossKind.Failure);
            Assert.Throws<InvalidDataException>(report.RequireNoLoss);
        }
    }
}
