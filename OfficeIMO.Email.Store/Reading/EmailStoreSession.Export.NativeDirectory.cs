using OfficeIMO.Drawing.Internal;
using OfficeIMO.Email;

namespace OfficeIMO.Email.Store;

public sealed partial class EmailStoreSession {
    /// <summary>Exports selected items to a Maildir or EMLX directory tree without modifying the source.</summary>
    public EmailStoreExportReport ExportToNativeDirectory(string destinationDirectory,
        EmailStoreNativeDirectoryExportOptions options, CancellationToken cancellationToken = default) {
        if (destinationDirectory == null) throw new ArgumentNullException(nameof(destinationDirectory));
        if (options == null) throw new ArgumentNullException(nameof(options));
        ThrowIfDisposed();
        string root = Path.GetFullPath(destinationDirectory);
        ThrowIfStoreSourceDestination(root, "Native directory export");
        Directory.CreateDirectory(root);
        var entries = new List<EmailStoreExportEntry>();
        var exportDiagnostics = new List<EmailStoreDiagnostic>();
        var paths = new EmailStoreExportPathBuilder(root, Folders, options.PreserveFolderHierarchy);
        var messageWriter = new EmailDocumentWriter(options.MessageOptions);
        var emlxWriter = new EmailStoreEmlxWriter(options.EmlxOptions);
        var enumeration = new EmailStoreEnumerationOptions(options.FolderId, options.IncludeDescendants,
            options.IncludeAssociatedItems, options.IncludeOrphanedItems,
            options.MaxItems == int.MaxValue ? int.MaxValue : options.MaxItems + 1);
        bool truncated = false;
        bool maildirFlagsCouldNotBeWritten = false;
        int attempted = 0;
        foreach (EmailStoreItemReference reference in EnumerateItems(enumeration, cancellationToken)) {
            cancellationToken.ThrowIfCancellationRequested();
            if (attempted >= options.MaxItems) { truncated = true; break; }
            attempted++;
            EmailStoreExportEntry entry = options.Format == EmailStoreNativeDirectoryFormat.Maildir
                ? ExportMaildirItem(reference, options, messageWriter, paths, ref maildirFlagsCouldNotBeWritten,
                    cancellationToken)
                : ExportEmlxItem(reference, options, emlxWriter, paths, cancellationToken);
            entries.Add(entry);
            if (!entry.Succeeded && !options.ContinueOnError) break;
        }

        if (maildirFlagsCouldNotBeWritten) {
            exportDiagnostics.Add(new EmailStoreDiagnostic(
                "EMAIL_STORE_MAILDIR_FLAGS_MANIFEST_ONLY",
                "The destination file system does not permit the Maildir ':2,' suffix; messages were delivered to new and flag state remains in the export manifest.",
                EmailStoreDiagnosticSeverity.Warning,
                root));
        }
        string? manifestPath = null;
        if (options.WriteManifest) {
            string candidate = Path.Combine(root, "officeimo-email-store-export.tsv");
            try {
                if (File.Exists(candidate) && !options.OverwriteExisting) {
                    exportDiagnostics.Add(new EmailStoreDiagnostic(
                        "EMAIL_STORE_EXPORT_MANIFEST_EXISTS",
                        "The export manifest already exists and overwriteExisting is false.",
                        EmailStoreDiagnosticSeverity.Warning, candidate));
                } else {
                    WriteExportManifest(candidate, root, entries, options.OverwriteExisting);
                    manifestPath = candidate;
                }
            } catch (Exception exception) when (exception is IOException || exception is UnauthorizedAccessException) {
                exportDiagnostics.Add(new EmailStoreDiagnostic("EMAIL_STORE_EXPORT_MANIFEST_FAILED",
                    exception.Message, EmailStoreDiagnosticSeverity.Error, candidate));
            }
        }
        return new EmailStoreExportReport(root, truncated, manifestPath, entries,
            Diagnostics.Concat(exportDiagnostics).ToArray());
    }

    private EmailStoreExportEntry ExportMaildirItem(EmailStoreItemReference reference,
        EmailStoreNativeDirectoryExportOptions options, EmailDocumentWriter writer,
        EmailStoreExportPathBuilder paths, ref bool flagsCouldNotBeWritten,
        CancellationToken cancellationToken) {
        var diagnostics = new List<EmailStoreDiagnostic>();
        string? destinationPath = null;
        long bytesWritten = 0;
        string? temporaryPath = null;
        string? maildirFlags = null;
        try {
            EmailStoreItem item = ReadItem(reference, cancellationToken);
            string folder = paths.GetItemDirectory(reference);
            string currentDirectory = Path.Combine(folder, "cur");
            string newDirectory = Path.Combine(folder, "new");
            string temporaryDirectory = Path.Combine(folder, "tmp");
            Directory.CreateDirectory(currentDirectory);
            Directory.CreateDirectory(newDirectory);
            Directory.CreateDirectory(temporaryDirectory);

            string flags = GetMaildirFlags(item.Document);
            maildirFlags = flags;
            bool supportsInfoSuffix = Array.IndexOf(Path.GetInvalidFileNameChars(), ':') < 0;
            bool useCurrent = supportsInfoSuffix;
            if (flags.Length > 0 && !supportsInfoSuffix) flagsCouldNotBeWritten = true;
            string fileName = BuildMaildirFileName(reference.Id);
            if (useCurrent) fileName += ":2," + flags;
            string finalDirectory = useCurrent ? currentDirectory : newDirectory;
            string candidate = Path.Combine(finalDirectory, fileName);
            if (File.Exists(candidate) && !options.OverwriteExisting)
                throw new IOException("The Maildir destination item already exists.");

            temporaryPath = Path.Combine(temporaryDirectory, fileName + "." + Guid.NewGuid().ToString("N") + ".tmp");
            EmailWriteResult result = writer.Write(item.Document, temporaryPath, EmailFileFormat.Eml);
            foreach (EmailDiagnostic diagnostic in result.Diagnostics)
                diagnostics.Add(ConvertDiagnostic(diagnostic, reference.Id));
            if (!result.HasErrors) {
                OfficeFileCommit.CommitTemporaryFile(temporaryPath, candidate,
                    options.OverwriteExisting ? OfficeFileCommit.ConflictPolicy.Replace :
                        OfficeFileCommit.ConflictPolicy.FailIfExists);
                temporaryPath = null;
                destinationPath = candidate;
                bytesWritten = result.BytesWritten;
            }
        } catch (Exception exception) when (exception is InvalidDataException || exception is NotSupportedException ||
                                             exception is IOException || exception is UnauthorizedAccessException ||
                                             exception is EmailStoreLimitExceededException ||
                                             exception is EmailLimitExceededException) {
            diagnostics.Add(new EmailStoreDiagnostic("EMAIL_STORE_MAILDIR_EXPORT_FAILED", exception.Message,
                EmailStoreDiagnosticSeverity.Error, "item/" + reference.Id));
        } finally {
            if (temporaryPath != null) OfficeFileCommit.DeleteIfExists(temporaryPath);
        }
        return new EmailStoreExportEntry(reference, destinationPath, bytesWritten, diagnostics, maildirFlags);
    }

    private EmailStoreExportEntry ExportEmlxItem(EmailStoreItemReference reference,
        EmailStoreNativeDirectoryExportOptions options, EmailStoreEmlxWriter writer,
        EmailStoreExportPathBuilder paths, CancellationToken cancellationToken) {
        var diagnostics = new List<EmailStoreDiagnostic>();
        string? destinationPath = null;
        long bytesWritten = 0;
        string? temporaryPath = null;
        try {
            EmailStoreItem item = ReadItem(reference, cancellationToken);
            string path = paths.GetItemPath(reference, item.Document.Subject, ".emlx");
            Directory.CreateDirectory(Path.GetDirectoryName(path)!);
            if (File.Exists(path) && !options.OverwriteExisting)
                throw new IOException("The EMLX destination item already exists.");
            temporaryPath = string.Concat(path, ".", Guid.NewGuid().ToString("N"), ".tmp");
            EmailWriteResult result;
            using (var stream = new FileStream(temporaryPath, FileMode.CreateNew, FileAccess.Write, FileShare.None)) {
                result = writer.Write(item.Document, stream);
                stream.Flush();
            }
            foreach (EmailDiagnostic diagnostic in result.Diagnostics)
                diagnostics.Add(ConvertDiagnostic(diagnostic, reference.Id));
            if (!result.HasErrors) {
                OfficeFileCommit.CommitTemporaryFile(temporaryPath, path,
                    options.OverwriteExisting ? OfficeFileCommit.ConflictPolicy.Replace :
                        OfficeFileCommit.ConflictPolicy.FailIfExists);
                temporaryPath = null;
                destinationPath = path;
                bytesWritten = result.BytesWritten;
            }
        } catch (Exception exception) when (exception is InvalidDataException || exception is NotSupportedException ||
                                             exception is IOException || exception is UnauthorizedAccessException ||
                                             exception is EmailStoreLimitExceededException ||
                                             exception is EmailLimitExceededException) {
            diagnostics.Add(new EmailStoreDiagnostic("EMAIL_STORE_EMLX_EXPORT_FAILED", exception.Message,
                EmailStoreDiagnosticSeverity.Error, "item/" + reference.Id));
        } finally {
            if (temporaryPath != null) OfficeFileCommit.DeleteIfExists(temporaryPath);
        }
        return new EmailStoreExportEntry(reference, destinationPath, bytesWritten, diagnostics);
    }

    private static string BuildMaildirFileName(string itemId) {
        string stable = EmailStoreExportPathBuilder.SanitizeSegment(itemId, 96, "item");
        return stable + "." + EmailStoreExportPathBuilder.GetStableHash(itemId) + ".officeimo";
    }

    private static string GetMaildirFlags(EmailDocument document) {
        var flags = new StringBuilder();
        if (document.MessageMetadata.IsDraft) flags.Append('D');
        if (PropertyFlag(document, "Emlx:Flag:Flagged")) flags.Append('F');
        if (PropertyFlag(document, "Emlx:Flag:Forwarded")) flags.Append('P');
        if (PropertyFlag(document, "Emlx:Flag:Answered")) flags.Append('R');
        if (document.MessageMetadata.IsRead == true) flags.Append('S');
        if (PropertyFlag(document, "Emlx:Flag:Deleted")) flags.Append('T');
        return flags.ToString();
    }

    private static bool PropertyFlag(EmailDocument document, string name) =>
        document.Properties.TryGetValue(name, out object? value) && value is bool enabled && enabled;
}
