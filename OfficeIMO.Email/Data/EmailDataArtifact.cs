using OfficeIMO.Email.AddressBook;
using OfficeIMO.Email.Store;

namespace OfficeIMO.Email.Data;

/// <summary>
/// Detects one persisted email-data artifact and dispatches it to OfficeIMO.Email, OfficeIMO.Email.Store, or
/// OfficeIMO.Email.AddressBook. It contains no alternate parsers or models.
/// </summary>
public static class EmailDataArtifact {
    private const int ContentLinePrefixBytes = 4096;

    /// <summary>Opens a supported file or directory with the existing artifact owner's bounded policy.</summary>
    public static EmailDataOpenResult Open(string path, EmailDataOpenOptions? options = null,
        CancellationToken cancellationToken = default) {
        if (path == null) throw new ArgumentNullException(nameof(path));
        cancellationToken.ThrowIfCancellationRequested();
        string fullPath = Path.GetFullPath(path);
        if (!File.Exists(fullPath) && !Directory.Exists(fullPath))
            throw new FileNotFoundException("The email-data artifact does not exist.", fullPath);

        EmailDataOpenOptions effective = options ?? EmailDataOpenOptions.Default;
        if (effective.ExpectedKind.HasValue)
            return OpenExpected(fullPath, effective.ExpectedKind.Value, effective, cancellationToken);

        if (Directory.Exists(fullPath))
            return OpenDirectory(fullPath, effective, cancellationToken);

        EmailStoreFormat storeFormat = EmailStoreReader.DetectFormat(fullPath);
        if (storeFormat != EmailStoreFormat.Unknown)
            return OpenStore(fullPath, effective, cancellationToken);

        if (string.Equals(Path.GetExtension(fullPath), ".oab", StringComparison.OrdinalIgnoreCase))
            return OpenAddressBook(fullPath, effective, cancellationToken);

        EmailDataArtifactKind contentLineKind = DetectContentLineKind(fullPath, cancellationToken);
        if (contentLineKind == EmailDataArtifactKind.Calendar ||
            string.Equals(Path.GetExtension(fullPath), ".ics", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(Path.GetExtension(fullPath), ".ical", StringComparison.OrdinalIgnoreCase))
            return OpenCalendar(fullPath, effective, cancellationToken);
        if (contentLineKind == EmailDataArtifactKind.Contact ||
            string.Equals(Path.GetExtension(fullPath), ".vcf", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(Path.GetExtension(fullPath), ".vcard", StringComparison.OrdinalIgnoreCase))
            return OpenContact(fullPath, effective, cancellationToken);

        return OpenEmail(fullPath, effective, cancellationToken);
    }

    private static EmailDataOpenResult OpenDirectory(string path, EmailDataOpenOptions options,
        CancellationToken cancellationToken) {
        OfflineAddressBookDiscoveryReport discovery;
        try {
            discovery = OfflineAddressBookInspector.Inspect(
                path, options.AddressBook, cancellationToken);
        } catch (OfflineAddressBookLimitExceededException exception) when (
            string.Equals(
                exception.LimitName,
                nameof(OfflineAddressBookReaderOptions.MaxDirectoryEntries),
                StringComparison.Ordinal)) {
            return OpenStore(path, options, cancellationToken);
        }
        if (discovery.ReadableFullDetailsCount > 0)
            return OpenAddressBook(path, options, cancellationToken);
        return OpenStore(path, options, cancellationToken);
    }

    private static EmailDataOpenResult OpenExpected(string path, EmailDataArtifactKind kind,
        EmailDataOpenOptions options, CancellationToken cancellationToken) {
        switch (kind) {
            case EmailDataArtifactKind.EmailDocument:
                return OpenEmail(path, options, cancellationToken);
            case EmailDataArtifactKind.Calendar:
                return OpenCalendar(path, options, cancellationToken);
            case EmailDataArtifactKind.Contact:
                return OpenContact(path, options, cancellationToken);
            case EmailDataArtifactKind.Store:
                return OpenStore(path, options, cancellationToken);
            case EmailDataArtifactKind.OfflineAddressBook:
                return OpenAddressBook(path, options, cancellationToken);
            default:
                throw new ArgumentOutOfRangeException(nameof(kind));
        }
    }

    private static EmailDataOpenResult OpenEmail(string path, EmailDataOpenOptions options,
        CancellationToken cancellationToken) {
        if (Directory.Exists(path))
            throw new InvalidDataException("An individual email artifact must be a file.");
        var reader = new EmailDocumentReader(options.Email);
        EmailReadResult result = options.UseStreamingEmailReader
            ? reader.ReadStreaming(path, cancellationToken)
            : reader.Read(path, cancellationToken);
        if (result.Document.Format != EmailFileFormat.Unknown) return new EmailDataOpenResult(path, result);
        string detail = result.Diagnostics.FirstOrDefault(diagnostic =>
            diagnostic.Severity == EmailDiagnosticSeverity.Error)?.Message ?? "The artifact format is unknown.";
        result.Dispose();
        throw new InvalidDataException(detail);
    }

    private static EmailDataOpenResult OpenCalendar(string path, EmailDataOpenOptions options,
        CancellationToken cancellationToken) {
        if (Directory.Exists(path)) throw new InvalidDataException("An iCalendar artifact must be a file.");
        return new EmailDataOpenResult(path,
            IcsDocument.Load(path, options.ContentLines, cancellationToken));
    }

    private static EmailDataOpenResult OpenContact(string path, EmailDataOpenOptions options,
        CancellationToken cancellationToken) {
        if (Directory.Exists(path)) throw new InvalidDataException("A vCard artifact must be a file.");
        return new EmailDataOpenResult(path,
            VCardDocument.Load(path, options.ContentLines, cancellationToken));
    }

    private static EmailDataOpenResult OpenStore(string path, EmailDataOpenOptions options,
        CancellationToken cancellationToken) => new EmailDataOpenResult(path,
            EmailStoreSession.Open(path, options.Store, cancellationToken));

    private static EmailDataOpenResult OpenAddressBook(string path, EmailDataOpenOptions options,
        CancellationToken cancellationToken) => new EmailDataOpenResult(path,
            OfflineAddressBookSession.Open(path, options.AddressBook, cancellationToken));

    private static EmailDataArtifactKind DetectContentLineKind(string path,
        CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        var bytes = new byte[ContentLinePrefixBytes];
        int count = 0;
        using (var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read,
                   ContentLinePrefixBytes, FileOptions.SequentialScan)) {
            while (count < bytes.Length) {
                cancellationToken.ThrowIfCancellationRequested();
                int read = stream.Read(bytes, count, bytes.Length - count);
                if (read == 0) break;
                count += read;
            }
        }
        string prefix = Encoding.UTF8.GetString(bytes, 0, count).TrimStart('\uFEFF', ' ', '\t', '\r', '\n');
        int lineEnd = prefix.IndexOfAny(new[] { '\r', '\n' });
        string firstLine = (lineEnd < 0 ? prefix : prefix.Substring(0, lineEnd)).Trim();
        if (string.Equals(firstLine, "BEGIN:VCALENDAR", StringComparison.OrdinalIgnoreCase))
            return EmailDataArtifactKind.Calendar;
        if (string.Equals(firstLine, "BEGIN:VCARD", StringComparison.OrdinalIgnoreCase))
            return EmailDataArtifactKind.Contact;
        return EmailDataArtifactKind.Unknown;
    }
}
