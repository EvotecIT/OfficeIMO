using OfficeIMO.Drawing.Internal;
using OfficeIMO.Email;

namespace OfficeIMO.Email.Store;

public sealed partial class EmailStoreSession {
    /// <summary>
    /// Streams selected messages through the OfficeIMO.Email mbox writer into a same-directory temporary file,
    /// then commits the completed mailbox. The source store is never modified.
    /// </summary>
    public EmailStoreMboxExportReport ExportToMbox(string destinationPath,
        EmailStoreMboxExportOptions? options = null,
        CancellationToken cancellationToken = default) {
        if (destinationPath == null) throw new ArgumentNullException(nameof(destinationPath));
        ThrowIfDisposed();
        EmailStoreMboxExportOptions effective = options ?? new EmailStoreMboxExportOptions();
        string destination = Path.GetFullPath(destinationPath);
        string directory = Path.GetDirectoryName(destination) ?? Path.GetFullPath(".");
        Directory.CreateDirectory(directory);
        var reportDiagnostics = new List<EmailStoreDiagnostic>();
        if (File.Exists(destination) && !effective.OverwriteExisting) {
            reportDiagnostics.Add(new EmailStoreDiagnostic(
                "EMAIL_STORE_EXPORT_DESTINATION_EXISTS",
                "The destination mailbox already exists and overwriteExisting is false.",
                EmailStoreDiagnosticSeverity.Error,
                destination));
            return new EmailStoreMboxExportReport(
                null, false, Array.Empty<EmailStoreMboxExportEntry>(),
                Diagnostics.Concat(reportDiagnostics).ToArray());
        }

        string temporary = Path.Combine(directory,
            string.Concat(".", Path.GetFileName(destination), ".", Guid.NewGuid().ToString("N"), ".tmp"));
        var entries = new List<EmailStoreMboxExportEntry>();
        bool truncated = false;
        bool commit = false;
        try {
            using (var output = new FileStream(
                temporary, FileMode.CreateNew, FileAccess.Write, FileShare.Read, 64 * 1024)) {
                var writer = new EmailMailboxWriter(effective.WriterOptions);
                int enumerationLimit = effective.MaxItems == int.MaxValue
                    ? int.MaxValue
                    : effective.MaxItems + 1;
                var enumeration = new EmailStoreEnumerationOptions(
                    effective.FolderId,
                    effective.IncludeDescendants,
                    effective.IncludeAssociatedItems,
                    effective.IncludeOrphanedItems,
                    enumerationLimit);
                int attempted = 0;
                foreach (EmailStoreItemReference reference in EnumerateItems(enumeration, cancellationToken)) {
                    cancellationToken.ThrowIfCancellationRequested();
                    if (attempted >= effective.MaxItems) {
                        truncated = true;
                        break;
                    }
                    attempted++;
                    EmailStoreMboxExportEntry entry = AppendMboxEntry(
                        reference, writer, output, cancellationToken);
                    entries.Add(entry);
                    if (!entry.Succeeded && !effective.ContinueOnError) break;
                }
                output.Flush();
            }

            if (entries.Count == 0 || entries.Any(item => item.Succeeded)) {
                CommitExportFile(temporary, destination, effective.OverwriteExisting);
                commit = true;
            } else {
                reportDiagnostics.Add(new EmailStoreDiagnostic(
                    "EMAIL_STORE_EXPORT_NO_ITEMS_WRITTEN",
                    "Every selected item failed, so no mailbox was committed.",
                    EmailStoreDiagnosticSeverity.Error,
                    destination));
            }
        } catch (Exception exception) when (
            exception is IOException || exception is UnauthorizedAccessException) {
            reportDiagnostics.Add(new EmailStoreDiagnostic(
                "EMAIL_STORE_EXPORT_COMMIT_FAILED",
                exception.Message,
                EmailStoreDiagnosticSeverity.Error,
                destination));
        } finally {
            if (!commit && File.Exists(temporary)) {
                try {
                    File.Delete(temporary);
                } catch (Exception exception) when (
                    exception is IOException || exception is UnauthorizedAccessException) {
                    reportDiagnostics.Add(new EmailStoreDiagnostic(
                        "EMAIL_STORE_EXPORT_TEMP_CLEANUP_FAILED",
                        exception.Message,
                        EmailStoreDiagnosticSeverity.Warning,
                        temporary));
                }
            }
        }

        return new EmailStoreMboxExportReport(
            commit ? destination : null,
            truncated,
            entries,
            Diagnostics.Concat(reportDiagnostics).ToArray());
    }

    /// <summary>
    /// Exports selected items one at a time through OfficeIMO.Email writers and records preservation diagnostics.
    /// The source store is never modified.
    /// </summary>
    public EmailStoreExportReport ExportToDirectory(string destinationDirectory,
        EmailStoreExportOptions? options = null,
        CancellationToken cancellationToken = default) {
        if (destinationDirectory == null) throw new ArgumentNullException(nameof(destinationDirectory));
        ThrowIfDisposed();
        EmailStoreExportOptions effective = options ?? new EmailStoreExportOptions();
        string root = Path.GetFullPath(destinationDirectory);
        Directory.CreateDirectory(root);
        var writer = new EmailDocumentWriter(effective.WriterOptions);
        var entries = new List<EmailStoreExportEntry>();
        var exportDiagnostics = new List<EmailStoreDiagnostic>();
        var paths = new EmailStoreExportPathBuilder(root, Folders, effective.PreserveFolderHierarchy);
        int enumerationLimit = effective.MaxItems == int.MaxValue
            ? int.MaxValue
            : effective.MaxItems + 1;
        var enumeration = new EmailStoreEnumerationOptions(
            effective.FolderId,
            effective.IncludeDescendants,
            effective.IncludeAssociatedItems,
            effective.IncludeOrphanedItems,
            enumerationLimit);
        bool truncated = false;
        int attempted = 0;
        foreach (EmailStoreItemReference reference in EnumerateItems(enumeration, cancellationToken)) {
            cancellationToken.ThrowIfCancellationRequested();
            if (attempted >= effective.MaxItems) {
                truncated = true;
                break;
            }
            attempted++;
            EmailStoreExportEntry entry = ExportItem(reference, effective, writer, paths, cancellationToken);
            entries.Add(entry);
            if (!entry.Succeeded && !effective.ContinueOnError) break;
        }

        string? manifestPath = null;
        if (effective.WriteManifest) {
            string candidate = Path.Combine(root, "officeimo-email-store-export.tsv");
            if (File.Exists(candidate) && !effective.OverwriteExisting) {
                exportDiagnostics.Add(new EmailStoreDiagnostic(
                    "EMAIL_STORE_EXPORT_MANIFEST_EXISTS",
                    "The export manifest already exists and overwriteExisting is false.",
                    EmailStoreDiagnosticSeverity.Warning,
                    candidate));
            } else {
                try {
                    WriteExportManifest(candidate, root, entries, effective.OverwriteExisting);
                    manifestPath = candidate;
                } catch (Exception exception) when (
                    exception is IOException || exception is UnauthorizedAccessException) {
                    exportDiagnostics.Add(new EmailStoreDiagnostic(
                        "EMAIL_STORE_EXPORT_MANIFEST_FAILED",
                        exception.Message,
                        EmailStoreDiagnosticSeverity.Error,
                        candidate));
                }
            }
        }

        return new EmailStoreExportReport(
            root,
            truncated,
            manifestPath,
            entries,
            Diagnostics.Concat(exportDiagnostics).ToArray());
    }

    private EmailStoreExportEntry ExportItem(EmailStoreItemReference reference,
        EmailStoreExportOptions options, EmailDocumentWriter writer,
        EmailStoreExportPathBuilder paths, CancellationToken cancellationToken) {
        var diagnostics = new List<EmailStoreDiagnostic>();
        string? destinationPath = null;
        long bytesWritten = 0;
        string? temporaryPath = null;
        try {
            EmailStoreItem item = ReadItem(reference, cancellationToken);
            string path = paths.GetItemPath(reference, item.Document.Subject, options.Format);
            Directory.CreateDirectory(Path.GetDirectoryName(path)!);
            if (File.Exists(path) && !options.OverwriteExisting) {
                diagnostics.Add(new EmailStoreDiagnostic(
                    "EMAIL_STORE_EXPORT_DESTINATION_EXISTS",
                    "The destination artifact already exists and overwriteExisting is false.",
                    EmailStoreDiagnosticSeverity.Error,
                    path));
                return new EmailStoreExportEntry(reference, null, 0, diagnostics);
            }

            temporaryPath = OfficeFileCommit.CreateStagingPath(path);
            EmailWriteResult result = writer.Write(item.Document, temporaryPath, options.Format);
            foreach (EmailDiagnostic diagnostic in result.Diagnostics) {
                diagnostics.Add(ConvertDiagnostic(diagnostic, reference.Id));
            }
            if (!result.HasErrors) {
                OfficeFileCommit.CommitTemporaryFile(temporaryPath, path,
                    options.OverwriteExisting ? OfficeFileCommit.ConflictPolicy.Replace :
                        OfficeFileCommit.ConflictPolicy.FailIfExists);
                temporaryPath = null;
                destinationPath = path;
                bytesWritten = result.BytesWritten;
            }
        } catch (EmailStoreLimitExceededException exception) {
            diagnostics.Add(new EmailStoreDiagnostic(
                "EMAIL_STORE_EXPORT_ITEM_LIMIT",
                exception.Message,
                EmailStoreDiagnosticSeverity.Error,
                string.Concat("item/", reference.Id)));
        } catch (Exception exception) when (
            exception is InvalidDataException ||
            exception is NotSupportedException ||
            exception is IOException ||
            exception is UnauthorizedAccessException) {
            diagnostics.Add(new EmailStoreDiagnostic(
                "EMAIL_STORE_EXPORT_ITEM_FAILED",
                exception.Message,
                EmailStoreDiagnosticSeverity.Error,
                string.Concat("item/", reference.Id)));
        } finally {
            OfficeFileCommit.DeleteIfExists(temporaryPath);
        }
        return new EmailStoreExportEntry(reference, destinationPath, bytesWritten, diagnostics);
    }

    private EmailStoreMboxExportEntry AppendMboxEntry(EmailStoreItemReference reference,
        EmailMailboxWriter writer, Stream output, CancellationToken cancellationToken) {
        var diagnostics = new List<EmailStoreDiagnostic>();
        long bytesWritten = 0;
        long initialLength = output.Length;
        try {
            EmailStoreItem item = ReadItem(reference, cancellationToken);
            var mailboxEntry = new EmailMailboxEntry(item.Document) {
                EnvelopeSender = item.Document.From?.Address,
                EnvelopeDate = item.Document.Date
            };
            EmailWriteResult result = writer.WriteEntries(
                new[] { mailboxEntry }, output, cancellationToken);
            foreach (EmailDiagnostic diagnostic in result.Diagnostics) {
                diagnostics.Add(ConvertDiagnostic(diagnostic, reference.Id));
            }
            if (!result.HasErrors) {
                bytesWritten = result.BytesWritten;
            } else {
                RollBackMboxEntry(output, initialLength);
            }
        } catch (EmailStoreLimitExceededException exception) {
            RollBackMboxEntry(output, initialLength);
            diagnostics.Add(new EmailStoreDiagnostic(
                "EMAIL_STORE_EXPORT_ITEM_LIMIT",
                exception.Message,
                EmailStoreDiagnosticSeverity.Error,
                string.Concat("item/", reference.Id)));
        } catch (Exception exception) when (
            exception is InvalidDataException ||
            exception is NotSupportedException ||
            exception is IOException) {
            RollBackMboxEntry(output, initialLength);
            diagnostics.Add(new EmailStoreDiagnostic(
                "EMAIL_STORE_EXPORT_ITEM_FAILED",
                exception.Message,
                EmailStoreDiagnosticSeverity.Error,
                string.Concat("item/", reference.Id)));
        }
        return new EmailStoreMboxExportEntry(reference, bytesWritten, diagnostics);
    }

    private static void RollBackMboxEntry(Stream output, long initialLength) {
        if (!output.CanSeek) return;
        output.SetLength(initialLength);
        output.Position = initialLength;
    }

    private static void CommitExportFile(string temporary, string destination, bool overwriteExisting) {
        if (!File.Exists(destination)) {
            File.Move(temporary, destination);
            return;
        }
        if (!overwriteExisting) {
            throw new IOException("The destination was created while the export was running.");
        }
        File.Replace(temporary, destination, destinationBackupFileName: null);
    }

    private static EmailStoreDiagnostic ConvertDiagnostic(EmailDiagnostic diagnostic, string itemId) =>
        new EmailStoreDiagnostic(
            diagnostic.Code,
            diagnostic.Message,
            diagnostic.Severity == EmailDiagnosticSeverity.Error
                ? EmailStoreDiagnosticSeverity.Error
                : diagnostic.Severity == EmailDiagnosticSeverity.Information
                    ? EmailStoreDiagnosticSeverity.Information
                    : EmailStoreDiagnosticSeverity.Warning,
            diagnostic.Location == null
                ? string.Concat("item/", itemId)
                : string.Concat("item/", itemId, "/", diagnostic.Location));

    private static void WriteExportManifest(string path, string root,
        IEnumerable<EmailStoreExportEntry> entries, bool overwriteExisting) {
        string temporaryPath = OfficeFileCommit.CreateStagingPath(path);
        try {
            OfficeFileCommit.EnsureTargetDirectory(path);
            using (var stream = new FileStream(
                temporaryPath, FileMode.CreateNew, FileAccess.Write, FileShare.Read))
            using (var writer = new StreamWriter(stream,
                new UTF8Encoding(encoderShouldEmitUTF8Identifier: false))) {
                writer.WriteLine("ItemId\tFolderId\tAssociated\tRecovered\tSucceeded\tBytes\tRelativePath\tMaildirFlags\tDiagnosticCodes");
                foreach (EmailStoreExportEntry entry in entries) {
                    string relativePath = entry.DestinationPath == null
                        ? string.Empty
                        : GetRelativePath(root, entry.DestinationPath);
                    writer.Write(EscapeManifest(entry.Reference.Id));
                    writer.Write('\t');
                    writer.Write(EscapeManifest(entry.Reference.FolderId));
                    writer.Write('\t');
                    writer.Write(entry.Reference.IsAssociated ? "true" : "false");
                    writer.Write('\t');
                    writer.Write(entry.Reference.IsOrphaned ? "true" : "false");
                    writer.Write('\t');
                    writer.Write(entry.Succeeded ? "true" : "false");
                    writer.Write('\t');
                    writer.Write(entry.BytesWritten.ToString(CultureInfo.InvariantCulture));
                    writer.Write('\t');
                    writer.Write(EscapeManifest(relativePath));
                    writer.Write('\t');
                    writer.Write(EscapeManifest(entry.MaildirFlags ?? string.Empty));
                    writer.Write('\t');
                    writer.Write(EscapeManifest(string.Join(",", entry.Diagnostics.Select(item => item.Code))));
                    writer.WriteLine();
                }
            }
            OfficeFileCommit.CommitTemporaryFile(temporaryPath, path,
                overwriteExisting ? OfficeFileCommit.ConflictPolicy.Replace :
                    OfficeFileCommit.ConflictPolicy.FailIfExists);
            temporaryPath = string.Empty;
        } finally {
            OfficeFileCommit.DeleteIfExists(temporaryPath);
        }
    }

    private static string GetRelativePath(string root, string path) {
        var rootUri = new Uri(AppendDirectorySeparator(root));
        var pathUri = new Uri(path);
        return Uri.UnescapeDataString(rootUri.MakeRelativeUri(pathUri).ToString())
            .Replace('/', Path.DirectorySeparatorChar);
    }

    private static string AppendDirectorySeparator(string path) =>
        path.EndsWith(Path.DirectorySeparatorChar.ToString(), StringComparison.Ordinal)
            ? path
            : string.Concat(path, Path.DirectorySeparatorChar.ToString());

    private static string EscapeManifest(string value) =>
        value.Replace('\t', ' ').Replace('\r', ' ').Replace('\n', ' ');
}
