using OfficeIMO.Email;

namespace OfficeIMO.Email.Store;

public sealed partial class EmailStoreSession {
    /// <summary>
    /// Discovers source-index orphans, exports each readable item independently, and records original item/folder
    /// evidence in a manifest. An isolated corrupt item does not invalidate successful artifacts when continuation
    /// is enabled. The source store is never modified.
    /// </summary>
    public EmailStoreRecoveryExportReport ExportRecoverableItemsToDirectory(
        string destinationDirectory,
        EmailStoreRecoveryExportOptions? options = null,
        CancellationToken cancellationToken = default) {
        if (destinationDirectory == null) throw new ArgumentNullException(nameof(destinationDirectory));
        ThrowIfDisposed();
        EmailStoreRecoveryExportOptions effective = options ??
            new EmailStoreRecoveryExportOptions();
        string root = Path.GetFullPath(destinationDirectory);
        ThrowIfStoreSourceDestination(root, "Recovery export");
        Directory.CreateDirectory(root);

        EmailStoreRecoveryReport discovery = DiscoverRecoverableItems(
            effective.DiscoveryOptions, cancellationToken);
        var exportOptions = new EmailStoreExportOptions(
            effective.Format,
            preserveFolderHierarchy: effective.PreserveFolderHierarchy,
            overwriteExisting: effective.OverwriteExisting,
            continueOnError: effective.ContinueOnItemError,
            writeManifest: false,
            maxItems: effective.DiscoveryOptions.MaxRecoveredItems,
            writerOptions: effective.WriterOptions);
        var writer = new EmailDocumentWriter(effective.WriterOptions);
        var paths = new EmailStoreExportPathBuilder(root, Folders, effective.PreserveFolderHierarchy);
        var entries = new List<EmailStoreExportEntry>(discovery.RecoveredItems.Count);
        foreach (EmailStoreItemReference reference in discovery.RecoveredItems) {
            cancellationToken.ThrowIfCancellationRequested();
            EmailStoreExportEntry entry = ExportItem(
                reference, exportOptions, writer, paths, cancellationToken);
            entries.Add(entry);
            if (!entry.Succeeded && !effective.ContinueOnItemError) break;
        }

        var diagnostics = new List<EmailStoreDiagnostic>();
        string? manifestPath = null;
        if (effective.WriteManifest) {
            string candidate = Path.Combine(root, "officeimo-email-store-recovery.tsv");
            if (File.Exists(candidate) && !effective.OverwriteExisting) {
                diagnostics.Add(new EmailStoreDiagnostic(
                    "EMAIL_STORE_RECOVERY_MANIFEST_EXISTS",
                    "The recovery manifest already exists and overwriteExisting is false.",
                    EmailStoreDiagnosticSeverity.Warning,
                    candidate));
            } else {
                try {
                    WriteExportManifest(candidate, root, entries, effective.OverwriteExisting);
                    manifestPath = candidate;
                } catch (Exception exception) when (
                    exception is IOException || exception is UnauthorizedAccessException) {
                    diagnostics.Add(new EmailStoreDiagnostic(
                        "EMAIL_STORE_RECOVERY_MANIFEST_FAILED",
                        exception.Message,
                        EmailStoreDiagnosticSeverity.Error,
                        candidate));
                }
            }
        }
        return new EmailStoreRecoveryExportReport(root, discovery, manifestPath,
            entries.AsReadOnly(), discovery.Diagnostics.Concat(diagnostics).ToArray());
    }
}
