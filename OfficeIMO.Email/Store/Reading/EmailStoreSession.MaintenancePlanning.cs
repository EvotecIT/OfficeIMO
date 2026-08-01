namespace OfficeIMO.Email.Store;

public sealed partial class EmailStoreSession {
    /// <summary>
    /// Produces a read-only, source-bound maintenance plan. Execution remains an explicit second step through the
    /// existing recovery export, PST compaction, or PST split plan APIs and always targets distinct paths.
    /// </summary>
    public EmailStoreMaintenancePlan PlanMaintenance(int maxItems = 100_000,
        CancellationToken cancellationToken = default) {
        if (maxItems <= 0) throw new ArgumentOutOfRangeException(nameof(maxItems));
        ThrowIfDisposed();
        string fingerprint = GetDurableSourceFingerprint(cancellationToken);
        var validation = Validate(new EmailStoreValidationOptions(
            mode: EmailStoreValidationMode.Summaries, maxItems: maxItems,
            includeAssociatedItems: true, includeOrphanedItems: true,
            verifyStructuralIntegrity: true), cancellationToken);
        var recovery = DiscoverRecoverableItems(new EmailStoreRecoveryOptions(
            maxItemsScanned: maxItems, maxRecoveredItems: maxItems), cancellationToken);
        var recommendations = new List<EmailStoreMaintenanceRecommendation>();
        if (recovery.RecoveredItems.Count > 0) {
            recommendations.Add(new EmailStoreMaintenanceRecommendation(
                EmailStoreMaintenanceAction.RecoveryExport,
                "The bounded source scan found items outside normal folder contents tables.",
                executableByOfficeImo: true, EmailDataLossRisk.None));
        }
        bool invalid = validation.ItemsFailed > 0 || validation.Diagnostics.Any(item =>
            item.Severity == EmailStoreDiagnosticSeverity.Error);
        if (invalid && (Format == EmailStoreFormat.Pst || Format == EmailStoreFormat.Ost)) {
            recommendations.Add(new EmailStoreMaintenanceRecommendation(
                EmailStoreMaintenanceAction.VerifiedRewrite,
                "Validation found source defects; plan a distinct-destination PST rewrite and require strict semantic verification.",
                executableByOfficeImo: true, EmailDataLossRisk.Possible));
        } else if (invalid) {
            recommendations.Add(new EmailStoreMaintenanceRecommendation(
                EmailStoreMaintenanceAction.ManualIntervention,
                "Validation found defects for a format without a native verified rewrite path.",
                executableByOfficeImo: false, EmailDataLossRisk.Unknown));
        }
        if (recommendations.Count == 0) {
            recommendations.Add(new EmailStoreMaintenanceRecommendation(
                EmailStoreMaintenanceAction.None, "No defect or orphan signal was found inside the configured bounds.",
                executableByOfficeImo: false, EmailDataLossRisk.None));
        }
        return new EmailStoreMaintenancePlan(fingerprint, Format, validation, recovery,
            recommendations.AsReadOnly());
    }
}
