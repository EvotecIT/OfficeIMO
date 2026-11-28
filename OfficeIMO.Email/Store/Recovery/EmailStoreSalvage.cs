using OfficeIMO.Core.Internal;

namespace OfficeIMO.Email.Store;

/// <summary>
/// Recovers readable items from damaged Outlook Express DBX databases and Outlook for Mac OLM archives that the
/// normal readers reject. Each item is written independently with a provenance manifest; the source is never
/// modified and nothing is written outside the destination directory.
/// </summary>
public static class EmailStoreSalvage {
    internal const string ManifestFileName = "officeimo-email-store-salvage.tsv";

    /// <summary>
    /// Scans a DBX message database for message records and linked MIME fragments, independent of its active
    /// index, and writes every readable message. Broken fragment chains yield partial messages with offsets.
    /// </summary>
    public static EmailStoreSalvageReport RecoverDbxToDirectory(string sourcePath, string destinationDirectory,
        EmailStoreSalvageOptions? options = null, CancellationToken cancellationToken = default) {
        ArgumentNullException.ThrowIfNull(sourcePath);
        using var source = new FileStream(sourcePath, FileMode.Open, FileAccess.Read, FileShare.Read,
            64 * 1024, FileOptions.RandomAccess);
        return RecoverCore(source, Path.GetFileName(sourcePath), Path.GetFullPath(sourcePath), destinationDirectory,
            options, EmailStoreFormat.Dbx, cancellationToken);
    }

    /// <summary>Scans a readable, seekable DBX stream. The stream position is restored afterwards.</summary>
    public static EmailStoreSalvageReport RecoverDbxToDirectory(Stream source, string? sourceName,
        string destinationDirectory, EmailStoreSalvageOptions? options = null,
        CancellationToken cancellationToken = default) =>
        RecoverCore(source, sourceName, (source as FileStream)?.Name, destinationDirectory, options,
            EmailStoreFormat.Dbx, cancellationToken);

    /// <summary>
    /// Reads an OLM archive through its central directory or, when that is damaged or missing, its local entry
    /// headers. Entries are verified against their CRC-32; unreadable XML entries are reported and skipped, and
    /// complete records in a truncated XML entry are still recovered.
    /// </summary>
    public static EmailStoreSalvageReport RecoverOlmToDirectory(string sourcePath, string destinationDirectory,
        EmailStoreSalvageOptions? options = null, CancellationToken cancellationToken = default) {
        ArgumentNullException.ThrowIfNull(sourcePath);
        using var source = new FileStream(sourcePath, FileMode.Open, FileAccess.Read, FileShare.Read,
            64 * 1024, FileOptions.RandomAccess);
        return RecoverCore(source, Path.GetFileName(sourcePath), Path.GetFullPath(sourcePath), destinationDirectory,
            options, EmailStoreFormat.Olm, cancellationToken);
    }

    /// <summary>
    /// Reads a readable, seekable OLM stream, which must not change during recovery; a change detected while copying
    /// verified data stops recovery with an <see cref="IOException"/>. The stream position is restored afterwards.
    /// </summary>
    public static EmailStoreSalvageReport RecoverOlmToDirectory(Stream source, string? sourceName,
        string destinationDirectory, EmailStoreSalvageOptions? options = null,
        CancellationToken cancellationToken = default) =>
        RecoverCore(source, sourceName, (source as FileStream)?.Name, destinationDirectory, options,
            EmailStoreFormat.Olm, cancellationToken);

    private static EmailStoreSalvageReport RecoverCore(Stream source, string? sourceName, string? sourcePath,
        string destinationDirectory, EmailStoreSalvageOptions? options, EmailStoreFormat format,
        CancellationToken cancellationToken) {
        ArgumentNullException.ThrowIfNull(source);
        ArgumentNullException.ThrowIfNull(destinationDirectory);
        if (!source.CanRead || !source.CanSeek)
            throw new ArgumentException("Salvage requires a readable, seekable source stream.", nameof(source));
        EmailStoreSalvageOptions effective = options ?? new EmailStoreSalvageOptions();
        string root = Path.GetFullPath(destinationDirectory);
        // The destination may hold the source file; each artifact path is checked against the source itself.
        if (sourcePath != null && EmailStorePathIdentity.IsSameOrDescendant(root, sourcePath))
            throw new InvalidOperationException("Salvage cannot use its source file as the destination directory.");
        long length = source.Length;
        if (length > effective.ReaderOptions.MaxInputBytes)
            throw new EmailStoreLimitExceededException(nameof(EmailStoreReaderOptions.MaxInputBytes),
                length, effective.ReaderOptions.MaxInputBytes);
        Directory.CreateDirectory(root);

        long originalPosition = source.Position;
        try {
            var output = new EmailStoreSalvageWriter(root, effective, sourcePath);
            EmailStoreSalvageReport report = format == EmailStoreFormat.Dbx
                ? new DbxSalvage(source, sourceName, effective, output, cancellationToken).Run()
                : new OlmSalvage(source, sourceName, effective, output, cancellationToken).Run();
            return report;
        } finally {
            source.Position = originalPosition;
        }
    }
}

/// <summary>Writes salvaged artifacts and the provenance manifest strictly inside one destination root.</summary>
internal sealed class EmailStoreSalvageWriter {
    private readonly EmailStoreSalvageOptions _options;
    private readonly EmailDocumentWriter _writer;

    private readonly string? _sourcePath;

    internal EmailStoreSalvageWriter(string root, EmailStoreSalvageOptions options, string? sourcePath) {
        Root = root;
        _sourcePath = sourcePath == null ? null : Path.GetFullPath(sourcePath);
        _options = options;
        _writer = new EmailDocumentWriter(options.WriterOptions);
    }

    internal string Root { get; }

    /// <summary>Creates a hidden spill file inside the destination root for raw or rebuilt source data.</summary>
    internal FileStream CreateSpill(string extension, bool deleteOnClose) {
        string path = Path.Combine(Root, ".officeimo-salvage-" + Guid.NewGuid().ToString("N") + extension);
        return new FileStream(path, FileMode.CreateNew, FileAccess.ReadWrite, FileShare.None, 64 * 1024,
            deleteOnClose ? FileOptions.DeleteOnClose : FileOptions.None);
    }

    internal string ResolvePath(EmailStoreExportPathBuilder paths, EmailStoreItemReference reference,
        string? subject, EmailFileFormat format) {
        string path = Path.GetFullPath(paths.GetItemPath(reference, subject, format));
        if (!EmailStorePathIdentity.IsSameOrDescendant(path, Root) ||
            string.Equals(path, Root, StringComparison.Ordinal))
            throw new InvalidOperationException("A salvage artifact path resolved outside the destination directory.");
        if (_sourcePath != null && EmailStorePathIdentity.AreEquivalent(path, _sourcePath))
            throw new IOException("A salvage artifact path would replace the source file.");
        return path;
    }

    /// <summary>Moves a completed spill file to its destination; returns false when the destination exists.</summary>
    internal bool CommitRaw(string spillPath, string path, string itemId, ICollection<EmailStoreDiagnostic> diagnostics) {
        Directory.CreateDirectory(Path.GetDirectoryName(path)!);
        if (File.Exists(path) && !_options.OverwriteExisting) {
            diagnostics.Add(DestinationExists(path, itemId));
            return false;
        }
        OfficeFileCommit.CommitTemporaryFile(spillPath, path, _options.OverwriteExisting
            ? OfficeFileCommit.ConflictPolicy.Replace : OfficeFileCommit.ConflictPolicy.FailIfExists);
        return true;
    }

    /// <summary>Serializes a projected document; returns the bytes written, or -1 when nothing was committed.</summary>
    internal long WriteDocument(EmailDocument document, string path, EmailFileFormat format, string itemId,
        ICollection<EmailStoreDiagnostic> diagnostics) {
        Directory.CreateDirectory(Path.GetDirectoryName(path)!);
        if (File.Exists(path) && !_options.OverwriteExisting) {
            diagnostics.Add(DestinationExists(path, itemId));
            return -1;
        }
        string? temporaryPath = OfficeFileCommit.CreateStagingPath(path);
        try {
            EmailWriteResult result;
            using (var stream = new FileStream(temporaryPath, FileMode.CreateNew, FileAccess.Write, FileShare.None)) {
                result = _writer.Write(document, stream, format);
                stream.Flush();
            }
            foreach (EmailDiagnostic diagnostic in result.Diagnostics) {
                // Disposition and loss risk travel with the diagnostic so the report can classify conversion loss.
                diagnostics.Add(new EmailStoreDiagnostic(diagnostic.Code, diagnostic.Message,
                    diagnostic.Severity == EmailDiagnosticSeverity.Error ? EmailStoreDiagnosticSeverity.Error
                    : diagnostic.Severity == EmailDiagnosticSeverity.Warning ? EmailStoreDiagnosticSeverity.Warning
                    : EmailStoreDiagnosticSeverity.Information,
                    diagnostic.Location == null ? itemId : itemId + "/" + diagnostic.Location,
                    diagnostic.Operation, diagnostic.ByteOffset, diagnostic.LimitName, diagnostic.ActualValue,
                    diagnostic.MaximumValue, diagnostic.Disposition, diagnostic.DataLossRisk, diagnostic.SuggestedAction,
                    diagnostic.IsRetryable));
            }
            if (result.HasErrors) return -1;
            OfficeFileCommit.CommitTemporaryFile(temporaryPath, path, _options.OverwriteExisting
                ? OfficeFileCommit.ConflictPolicy.Replace : OfficeFileCommit.ConflictPolicy.FailIfExists);
            temporaryPath = null;
            return result.BytesWritten;
        } finally {
            OfficeFileCommit.DeleteIfExists(temporaryPath);
        }
    }

    internal string? WriteManifest(IReadOnlyList<EmailStoreSalvageEntry> entries,
        ICollection<EmailStoreDiagnostic> diagnostics) {
        if (!_options.WriteManifest) return null;
        string path = Path.Combine(Root, EmailStoreSalvage.ManifestFileName);
        if (_sourcePath != null && EmailStorePathIdentity.AreEquivalent(path, _sourcePath)) {
            diagnostics.Add(new EmailStoreDiagnostic("EMAIL_STORE_SALVAGE_MANIFEST_IS_SOURCE",
                "The salvage manifest path is the source file, so no manifest was written.",
                EmailStoreDiagnosticSeverity.Error, path));
            return null;
        }
        if (File.Exists(path) && !_options.OverwriteExisting) {
            diagnostics.Add(new EmailStoreDiagnostic("EMAIL_STORE_SALVAGE_MANIFEST_EXISTS",
                "The salvage manifest already exists and overwriteExisting is false.",
                EmailStoreDiagnosticSeverity.Warning, path));
            return null;
        }
        string? temporaryPath = OfficeFileCommit.CreateStagingPath(path);
        try {
            using (var stream = new FileStream(temporaryPath, FileMode.CreateNew, FileAccess.Write, FileShare.Read))
            using (var writer = new StreamWriter(stream, new UTF8Encoding(encoderShouldEmitUTF8Identifier: false))) {
                writer.WriteLine("ItemId\tStatus\tKind\tSourceLocation\tSourceOffset\tRecoveredBytes\tDeclaredBytes\tInActiveIndex\tDamageOffset\tDamageCode\tBytesWritten\tRelativePath\tDiagnosticCount");
                foreach (EmailStoreSalvageEntry entry in entries) {
                    writer.WriteLine(string.Join("\t",
                        Escape(entry.ItemId),
                        entry.Status.ToString(),
                        entry.Kind?.ToString() ?? string.Empty,
                        Escape(entry.SourceLocation),
                        Offset(entry.SourceOffset),
                        entry.RecoveredBytes.ToString(CultureInfo.InvariantCulture),
                        entry.DeclaredBytes?.ToString(CultureInfo.InvariantCulture) ?? string.Empty,
                        entry.IsInActiveIndex.HasValue ? (entry.IsInActiveIndex.Value ? "true" : "false") : string.Empty,
                        Offset(entry.DamageOffset),
                        entry.DamageCode ?? string.Empty,
                        entry.BytesWritten.ToString(CultureInfo.InvariantCulture),
                        Escape(entry.DestinationPath == null ? string.Empty : RelativePath(entry.DestinationPath)),
                        entry.Diagnostics.Count.ToString(CultureInfo.InvariantCulture)));
                }
            }
            OfficeFileCommit.CommitTemporaryFile(temporaryPath, path, _options.OverwriteExisting
                ? OfficeFileCommit.ConflictPolicy.Replace : OfficeFileCommit.ConflictPolicy.FailIfExists);
            temporaryPath = null;
            return path;
        } catch (Exception exception) when (exception is IOException || exception is UnauthorizedAccessException) {
            diagnostics.Add(new EmailStoreDiagnostic("EMAIL_STORE_SALVAGE_MANIFEST_FAILED", exception.Message,
                EmailStoreDiagnosticSeverity.Error, path));
            return null;
        } finally {
            OfficeFileCommit.DeleteIfExists(temporaryPath);
        }
    }

    private string RelativePath(string path) =>
        path.Substring(Root.Length).TrimStart(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);

    private static string Offset(long? value) =>
        value.HasValue ? "0x" + value.Value.ToString("x", CultureInfo.InvariantCulture) : string.Empty;

    private static string Escape(string value) => value.Replace('\t', ' ').Replace('\r', ' ').Replace('\n', ' ');

    private static EmailStoreDiagnostic DestinationExists(string path, string itemId) =>
        new EmailStoreDiagnostic("EMAIL_STORE_SALVAGE_DESTINATION_EXISTS",
            "The destination artifact already exists and overwriteExisting is false.",
            EmailStoreDiagnosticSeverity.Error, itemId + " -> " + path);
}
