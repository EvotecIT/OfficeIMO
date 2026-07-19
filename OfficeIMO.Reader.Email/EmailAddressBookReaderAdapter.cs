using OfficeIMO.Email;

namespace OfficeIMO.Reader.Email;

internal static class EmailAddressBookReaderAdapter {
    internal static IEnumerable<ReaderChunk> Read(
        string path, ReaderOptions readerOptions,
        ReaderEmailAddressBookOptions adapterOptions, CancellationToken cancellationToken) =>
        EmailAddressBookEntryReader.Read(path, readerOptions, adapterOptions, cancellationToken)
            .SelectMany(result => result.Chunks)
            .ToArray();

    internal static IEnumerable<ReaderChunk> Read(
        Stream stream, string? sourceName, ReaderOptions readerOptions,
        ReaderEmailAddressBookOptions adapterOptions, CancellationToken cancellationToken) =>
        EmailAddressBookEntryReader.Read(stream, NormalizeSourceName(sourceName),
                readerOptions, adapterOptions, cancellationToken)
            .SelectMany(result => result.Chunks)
            .ToArray();

    internal static OfficeDocumentReadResult ReadDocument(
        string path, ReaderOptions readerOptions,
        ReaderEmailAddressBookOptions adapterOptions, CancellationToken cancellationToken) {
        ReaderEmailAddressBookEntryResult[] entries = EmailAddressBookEntryReader.Read(
            path, readerOptions, adapterOptions, cancellationToken).ToArray();
        string? sourceHash = adapterOptions.ComputeSourceHash ? TryHashFile(path) : null;
        var source = new OfficeDocumentSource {
            Path = path,
            SourceId = entries.SelectMany(entry => entry.Chunks).FirstOrDefault()?.SourceId,
            SourceHash = sourceHash,
            LengthBytes = TryGetLength(path),
            LastWriteUtc = TryGetLastWrite(path)
        };
        return CreateDocumentResult(entries, source);
    }

    internal static OfficeDocumentReadResult ReadDocument(
        Stream stream, string? sourceName, ReaderOptions readerOptions,
        ReaderEmailAddressBookOptions adapterOptions, CancellationToken cancellationToken) {
        string logicalName = NormalizeSourceName(sourceName);
        ReaderEmailAddressBookEntryResult[] entries = EmailAddressBookEntryReader.Read(
            stream, logicalName, readerOptions, adapterOptions, cancellationToken).ToArray();
        string? sourceHash = adapterOptions.ComputeSourceHash ? TryHashStream(stream) : null;
        var source = new OfficeDocumentSource {
            Path = logicalName,
            SourceId = entries.SelectMany(entry => entry.Chunks).FirstOrDefault()?.SourceId,
            SourceHash = sourceHash,
            LengthBytes = TryGetLength(stream)
        };
        return CreateDocumentResult(entries, source);
    }

    private static OfficeDocumentReadResult CreateDocumentResult(
        IReadOnlyList<ReaderEmailAddressBookEntryResult> entries,
        OfficeDocumentSource source) {
        ReaderChunk[] chunks = entries.SelectMany(entry => entry.Chunks).ToArray();
        OfficeDocumentReadResult result = DocumentReaderEngine.CreateDocumentResult(
            chunks,
            ReaderInputKind.Email,
            source,
            new[] { OfficeDocumentReaderBuilderEmailAddressBookExtensions.HandlerId });
        result.Diagnostics = entries.SelectMany(entry => entry.Diagnostics)
            .Select(EmailAddressBookReaderProjection.MapDiagnostic)
            .ToArray();
        result.Metadata = new[] {
            new OfficeDocumentMetadataEntry {
                Id = "oab-projected-entry-count",
                Category = "address-book",
                Name = "ProjectedEntryCount",
                Value = entries.Count.ToString(CultureInfo.InvariantCulture),
                ValueType = "number"
            },
            new OfficeDocumentMetadataEntry {
                Id = "oab-failed-entry-count",
                Category = "address-book",
                Name = "FailedEntryCount",
                Value = entries.Count(entry => !entry.Succeeded).ToString(CultureInfo.InvariantCulture),
                ValueType = "number"
            }
        };
        return result;
    }

    private static string NormalizeSourceName(string? sourceName) =>
        string.IsNullOrWhiteSpace(sourceName) ? "address-book.oab" : sourceName!.Trim();

    private static long? TryGetLength(string path) {
        try {
            return File.Exists(path) ? new FileInfo(path).Length : (long?)null;
        } catch {
            return null;
        }
    }

    private static DateTime? TryGetLastWrite(string path) {
        try {
            return File.Exists(path) ? new FileInfo(path).LastWriteTimeUtc : (DateTime?)null;
        } catch {
            return null;
        }
    }

    private static long? TryGetLength(Stream stream) {
        try {
            return stream.CanSeek ? stream.Length - stream.Position : (long?)null;
        } catch {
            return null;
        }
    }

    private static string? TryHashFile(string path) {
        try {
            using (var stream = new FileStream(path, FileMode.Open, FileAccess.Read,
                FileShare.ReadWrite | FileShare.Delete)) {
                return Hash(stream);
            }
        } catch {
            return null;
        }
    }

    private static string? TryHashStream(Stream stream) {
        if (!stream.CanSeek) return null;
        long position = stream.Position;
        try {
            return Hash(stream);
        } catch {
            return null;
        } finally {
            stream.Position = position;
        }
    }

    private static string Hash(Stream stream) {
        using (SHA256 sha = SHA256.Create()) {
            byte[] hash = sha.ComputeHash(stream);
            var result = new StringBuilder(hash.Length * 2);
            foreach (byte item in hash) result.Append(item.ToString("x2", CultureInfo.InvariantCulture));
            return result.ToString();
        }
    }
}
