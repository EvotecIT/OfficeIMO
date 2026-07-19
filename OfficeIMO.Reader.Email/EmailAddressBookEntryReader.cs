using OfficeIMO.Email;
using OfficeIMO.Email.AddressBook;

namespace OfficeIMO.Reader.Email;

/// <summary>Item-at-a-time Reader ingestion for Outlook Offline Address Book entries.</summary>
public static class EmailAddressBookEntryReader {
    /// <summary>Lazily projects selected OAB entries without hashing or retaining the complete source.</summary>
    public static IEnumerable<ReaderEmailAddressBookEntryResult> Read(
        string path,
        ReaderOptions? readerOptions = null,
        ReaderEmailAddressBookOptions? addressBookOptions = null,
        CancellationToken cancellationToken = default) =>
        Read(OfficeDocumentReader.Default, path, readerOptions, addressBookOptions, cancellationToken);

    /// <summary>Lazily projects selected OAB entries in the supplied Reader scope.</summary>
    public static IEnumerable<ReaderEmailAddressBookEntryResult> Read(
        OfficeDocumentReader reader,
        string path,
        ReaderOptions? readerOptions = null,
        ReaderEmailAddressBookOptions? addressBookOptions = null,
        CancellationToken cancellationToken = default) {
        if (reader == null) throw new ArgumentNullException(nameof(reader));
        if (path == null) throw new ArgumentNullException(nameof(path));
        return reader.Scope(ReadPathCore(path, readerOptions ?? new ReaderOptions(),
            ReaderEmailAddressBookOptionsCloner.CloneOrDefault(addressBookOptions), cancellationToken));
    }

    /// <summary>
    /// Lazily projects a seekable OAB Full Details stream. Non-seekable streams are bounded and buffered through
    /// the Reader input-limit service. The caller-owned stream remains open.
    /// </summary>
    public static IEnumerable<ReaderEmailAddressBookEntryResult> Read(
        Stream stream,
        string sourceName,
        ReaderOptions? readerOptions = null,
        ReaderEmailAddressBookOptions? addressBookOptions = null,
        CancellationToken cancellationToken = default) {
        if (stream == null) throw new ArgumentNullException(nameof(stream));
        if (sourceName == null) throw new ArgumentNullException(nameof(sourceName));
        return ReadStreamCore(stream, sourceName, readerOptions ?? new ReaderOptions(),
            ReaderEmailAddressBookOptionsCloner.CloneOrDefault(addressBookOptions), cancellationToken);
    }

    private static IEnumerable<ReaderEmailAddressBookEntryResult> ReadPathCore(
        string path, ReaderOptions readerOptions, ReaderEmailAddressBookOptions adapterOptions,
        CancellationToken cancellationToken) {
        OfflineAddressBookReaderOptions effective =
            ReaderEmailAddressBookOptionsCloner.CreateEffective(adapterOptions, readerOptions);
        using (OfflineAddressBookSession session = OfflineAddressBookSession.Open(
            path, effective, cancellationToken)) {
            foreach (ReaderEmailAddressBookEntryResult result in ReadSession(
                session, path, readerOptions, adapterOptions, cancellationToken)) {
                yield return result;
            }
        }
    }

    private static IEnumerable<ReaderEmailAddressBookEntryResult> ReadStreamCore(
        Stream stream, string sourceName, ReaderOptions readerOptions,
        ReaderEmailAddressBookOptions adapterOptions, CancellationToken cancellationToken) {
        OfflineAddressBookReaderOptions effective =
            ReaderEmailAddressBookOptionsCloner.CreateEffective(adapterOptions, readerOptions);
        Stream parseStream = ReaderInputLimits.EnsureSeekableReadStream(
            stream, effective.MaxInputBytes, cancellationToken, out bool ownsParseStream);
        try {
            using (OfflineAddressBookSession session = OfflineAddressBookSession.Open(
                parseStream, sourceName, effective, cancellationToken)) {
                foreach (ReaderEmailAddressBookEntryResult result in ReadSession(
                    session, sourceName, readerOptions, adapterOptions, cancellationToken)) {
                    yield return result;
                }
            }
        } finally {
            if (ownsParseStream) parseStream.Dispose();
        }
    }

    private static IEnumerable<ReaderEmailAddressBookEntryResult> ReadSession(
        OfflineAddressBookSession session,
        string sourceName,
        ReaderOptions readerOptions,
        ReaderEmailAddressBookOptions adapterOptions,
        CancellationToken cancellationToken) {
        EmailDiagnostic[] openingDiagnostics = session.Diagnostics.ToArray();
        IEnumerable<OfflineAddressBookEntryReference> references;
        IReadOnlyDictionary<string, OfflineAddressBookEntrySummary> searchSummaries =
            new Dictionary<string, OfflineAddressBookEntrySummary>(StringComparer.Ordinal);
        if (adapterOptions.Query == null) {
            references = session.EnumerateEntryReferences(
                new OfflineAddressBookEnumerationOptions(
                    adapterOptions.AddressListId,
                    adapterOptions.MaxEntries,
                    adapterOptions.ContinueOnEntryError),
                cancellationToken);
        } else {
            OfflineAddressBookSearchReport report = session.Search(
                ReaderEmailAddressBookOptionsCloner.CreateEffectiveQuery(
                    adapterOptions.Query, adapterOptions),
                cancellationToken: cancellationToken);
            openingDiagnostics = openingDiagnostics.Concat(report.Diagnostics).ToArray();
            OfflineAddressBookSearchResult[] matches = report.Results
                .Take(adapterOptions.MaxEntries)
                .ToArray();
            searchSummaries = matches.ToDictionary(
                result => result.Summary.Reference.Id,
                result => result.Summary,
                StringComparer.Ordinal);
            references = matches.Select(result => result.Summary.Reference);
        }

        int index = 0;
        foreach (OfflineAddressBookEntryReference reference in references) {
            cancellationToken.ThrowIfCancellationRequested();
            if (index >= adapterOptions.MaxEntries) yield break;
            string logicalPath = EmailAddressBookReaderProjection.LogicalPath(sourceName, reference);
            OfflineAddressBookEntrySummary? summary = searchSummaries.TryGetValue(reference.Id, out OfflineAddressBookEntrySummary? found)
                ? found
                : null;
            ReaderEmailAddressBookEntryResult result;
            try {
                OfflineAddressBookEntry entry = session.ReadEntry(reference, cancellationToken);
                summary = summary ?? entry.ToSummary();
                ReaderChunk chunk = EmailAddressBookReaderProjection.CreateChunk(
                    entry, logicalPath, readerOptions, adapterOptions);
                result = new ReaderEmailAddressBookEntryResult(
                    reference, summary, logicalPath, new[] { chunk }, entry.Diagnostics,
                    index == 0 ? openingDiagnostics : null);
            } catch (Exception exception) when (
                adapterOptions.ContinueOnEntryError &&
                (exception is InvalidDataException ||
                 exception is NotSupportedException ||
                 exception is KeyNotFoundException ||
                 exception is OfflineAddressBookLimitExceededException)) {
                EmailDiagnostic diagnostic = EmailAddressBookReaderProjection.ProjectionError(
                    exception, reference.Id);
                result = new ReaderEmailAddressBookEntryResult(
                    reference, summary, logicalPath, Array.Empty<ReaderChunk>(),
                    new[] { diagnostic }, index == 0 ? openingDiagnostics : null);
            }
            index++;
            yield return result;
        }
    }
}
