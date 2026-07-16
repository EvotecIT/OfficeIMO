using OfficeIMO.Email;

namespace OfficeIMO.Email.AddressBook;

/// <summary>
/// Read-only, lazy session over one OAB v4 Full Details file or a directory containing multiple address lists.
/// Sessions are not thread-safe.
/// </summary>
public sealed partial class OfflineAddressBookSession : IDisposable {
    private readonly OfflineAddressBookReaderOptions _options;
    private readonly List<OabAddressListSource> _sources;
    private readonly List<EmailDiagnostic> _diagnostics;
    private bool _disposed;

    private OfflineAddressBookSession(string sourcePath,
        OfflineAddressBookReaderOptions options,
        IReadOnlyList<OfflineAddressBookFileInfo> files,
        List<OabAddressListSource> sources,
        List<EmailDiagnostic> diagnostics) {
        SourcePath = sourcePath;
        _options = options;
        Files = files;
        _sources = sources;
        _diagnostics = diagnostics;
        AddressLists = sources.Select(source => source.Info).ToArray();
        long total = 0;
        foreach (OabAddressListSource source in sources) total = checked(total + source.Info.DeclaredEntryCount);
        DeclaredEntryCount = total;
    }

    /// <summary>Path or stream name used to open the session.</summary>
    public string SourcePath { get; }
    /// <summary>All discovered OAB components, including supporting legacy indexes and templates.</summary>
    public IReadOnlyList<OfflineAddressBookFileInfo> Files { get; }
    /// <summary>Modern OAB v4 Full Details address lists available for entry enumeration.</summary>
    public IReadOnlyList<OfflineAddressBookListInfo> AddressLists { get; }
    /// <summary>Total declared records across all address lists.</summary>
    public long DeclaredEntryCount { get; }
    /// <summary>Open, discovery, and recoverable entry diagnostics.</summary>
    public IReadOnlyList<EmailDiagnostic> Diagnostics => _diagnostics;

    /// <summary>Opens one Full Details file or discovers every Full Details component below a directory.</summary>
    public static OfflineAddressBookSession Open(string path,
        OfflineAddressBookReaderOptions? options = null,
        CancellationToken cancellationToken = default) {
        if (path == null) throw new ArgumentNullException(nameof(path));
        OfflineAddressBookReaderOptions effective = options ?? OfflineAddressBookReaderOptions.Default;
        bool isSingleFile = File.Exists(Path.GetFullPath(path));
        OabDiscoveryResult discovery = OabFileDiscovery.Discover(path, effective, cancellationToken);
        if (discovery.FullDetailsSources.Count == 0) {
            string detail = discovery.Files.Count == 0
                ? "No .oab components were discovered."
                : "No uncompressed OAB v4 Full Details component was discovered.";
            throw new NotSupportedException(detail);
        }

        var diagnostics = new List<EmailDiagnostic>(discovery.Diagnostics);
        var sources = new List<OabAddressListSource>();
        foreach (OabSource source in discovery.FullDetailsSources) {
            cancellationToken.ThrowIfCancellationRequested();
            try {
                OfflineAddressBookListInfo info = OabV4MetadataReader.Read(
                    source, sources.Count, effective, diagnostics);
                sources.Add(new OabAddressListSource(source, info));
            } catch (Exception exception) when (!isSingleFile && IsRecoverableReadException(exception)) {
                diagnostics.Add(new EmailDiagnostic(
                    "OAB_FULL_DETAILS_SKIPPED",
                    exception.Message,
                    EmailDiagnosticSeverity.Error,
                    source.SourcePath));
            }
        }
        if (sources.Count == 0) throw new InvalidDataException("No readable OAB v4 Full Details component was found.");
        return new OfflineAddressBookSession(Path.GetFullPath(path), effective,
            discovery.Files, sources, diagnostics);
    }

    /// <summary>
    /// Opens an uncompressed OAB v4 Full Details stream without taking ownership. The stream must be seekable.
    /// Its original position is restored after each operation and when the session is disposed.
    /// </summary>
    public static OfflineAddressBookSession Open(Stream stream, string sourceName,
        OfflineAddressBookReaderOptions? options = null,
        CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        OfflineAddressBookReaderOptions effective = options ?? OfflineAddressBookReaderOptions.Default;
        OabSource source = OabSource.FromStream(stream, sourceName);
        if (source.Length > effective.MaxInputBytes) {
            throw new OfflineAddressBookLimitExceededException(
                nameof(effective.MaxInputBytes), source.Length, effective.MaxInputBytes, sourceName);
        }
        OfflineAddressBookFileInfo file = OabFileDiscovery.InspectStream(source);
        if (file.Format != OfflineAddressBookFormat.Version4FullDetails) {
            throw new NotSupportedException(string.Concat(
                "The stream is not an uncompressed OAB v4 Full Details component (version 0x",
                file.Version.ToString("X8", CultureInfo.InvariantCulture), ")."));
        }
        var diagnostics = new List<EmailDiagnostic>();
        OfflineAddressBookListInfo info = OabV4MetadataReader.Read(source, 0, effective, diagnostics);
        return new OfflineAddressBookSession(sourceName, effective, new[] { file },
            new List<OabAddressListSource> { new OabAddressListSource(source, info) }, diagnostics);
    }

    /// <summary>Lazily enumerates record references without decoding property values.</summary>
    public IEnumerable<OfflineAddressBookEntryReference> EnumerateEntryReferences(
        OfflineAddressBookEnumerationOptions? options = null,
        CancellationToken cancellationToken = default) {
        ThrowIfDisposed();
        return EnumerateEntryReferencesCore(options ?? new OfflineAddressBookEnumerationOptions(), cancellationToken);
    }

    /// <summary>Lazily decodes entries, keeping memory related to the active record rather than total OAB size.</summary>
    public IEnumerable<OfflineAddressBookEntry> EnumerateEntries(
        OfflineAddressBookEnumerationOptions? options = null,
        CancellationToken cancellationToken = default) {
        ThrowIfDisposed();
        return EnumerateEntriesCore(options ?? new OfflineAddressBookEnumerationOptions(), cancellationToken);
    }

    /// <summary>Random-access reads one previously enumerated record reference.</summary>
    public OfflineAddressBookEntry ReadEntry(OfflineAddressBookEntryReference reference,
        CancellationToken cancellationToken = default) {
        ThrowIfDisposed();
        if (reference == null) throw new ArgumentNullException(nameof(reference));
        cancellationToken.ThrowIfCancellationRequested();
        OabAddressListSource selected = ValidateReference(reference);
        using (OabStreamLease lease = selected.Source.OpenRead()) {
            Stream stream = lease.Stream;
            OabBinary.Seek(selected.Source, stream, reference.RecordOffset, reference.Id);
            OabRecordEnvelope envelope = OabV4RecordReader.ReadEnvelope(
                selected.Source, stream, _options, reference.Id);
            if (envelope.Size != reference.RecordLength) {
                throw new InvalidDataException("OAB record length changed after the reference was enumerated.");
            }
            OabParsedRecord record = OabV4RecordReader.Parse(
                envelope, selected.Info.EntryPropertyDefinitions, _options, reference.Id);
            return new OfflineAddressBookEntry(reference, selected.Info, record.Properties, record.Diagnostics);
        }
    }

    /// <summary>Restores caller-owned stream positions. Source files are opened only during individual operations.</summary>
    public void Dispose() {
        if (_disposed) return;
        _disposed = true;
        foreach (OabAddressListSource source in _sources) source.Source.RestoreCallerPosition();
    }

    private IEnumerable<OfflineAddressBookEntryReference> EnumerateEntryReferencesCore(
        OfflineAddressBookEnumerationOptions options,
        CancellationToken cancellationToken) {
        int returned = 0;
        foreach (OabAddressListSource selected in SelectSources(options.AddressListId)) {
            using (OabStreamLease lease = selected.Source.OpenRead()) {
                Stream stream = lease.Stream;
                OabBinary.Seek(selected.Source, stream, selected.Info.EntriesOffset, selected.Info.Id);
                for (long entryIndex = 0; entryIndex < selected.Info.DeclaredEntryCount; entryIndex++) {
                    cancellationToken.ThrowIfCancellationRequested();
                    if (returned >= options.MaxEntries) yield break;
                    long offset = stream.Position - selected.Source.BaseOffset;
                    string location = BuildEntryLocation(selected.Info, entryIndex);
                    int size;
                    try {
                        size = OabV4RecordReader.ReadRecordSize(selected.Source, stream, _options, location);
                    } catch (Exception exception) when (options.ContinueOnEntryError && IsRecoverableReadException(exception)) {
                        AddFramingDiagnostic(exception, location);
                        break;
                    }
                    stream.Position = checked(selected.Source.BaseOffset + offset + size);
                    returned++;
                    yield return new OfflineAddressBookEntryReference(
                        selected.Info.Id, selected.Info.Index, entryIndex, offset, size,
                        selected.SnapshotId);
                }
            }
        }
    }

    private IEnumerable<OfflineAddressBookEntry> EnumerateEntriesCore(
        OfflineAddressBookEnumerationOptions options,
        CancellationToken cancellationToken) {
        int returned = 0;
        foreach (OabAddressListSource selected in SelectSources(options.AddressListId)) {
            using (OabStreamLease lease = selected.Source.OpenRead()) {
                Stream stream = lease.Stream;
                OabBinary.Seek(selected.Source, stream, selected.Info.EntriesOffset, selected.Info.Id);
                for (long entryIndex = 0; entryIndex < selected.Info.DeclaredEntryCount; entryIndex++) {
                    cancellationToken.ThrowIfCancellationRequested();
                    if (returned >= options.MaxEntries) yield break;
                    long offset = stream.Position - selected.Source.BaseOffset;
                    string location = BuildEntryLocation(selected.Info, entryIndex);
                    OabRecordEnvelope envelope;
                    try {
                        envelope = OabV4RecordReader.ReadEnvelope(selected.Source, stream, _options, location);
                    } catch (Exception exception) when (options.ContinueOnEntryError && IsRecoverableReadException(exception)) {
                        AddFramingDiagnostic(exception, location);
                        break;
                    }
                    var reference = new OfflineAddressBookEntryReference(
                        selected.Info.Id, selected.Info.Index, entryIndex, offset, envelope.Size,
                        selected.SnapshotId);
                    OabParsedRecord record;
                    try {
                        record = OabV4RecordReader.Parse(
                            envelope, selected.Info.EntryPropertyDefinitions, _options, reference.Id);
                    } catch (Exception exception) when (options.ContinueOnEntryError && IsRecoverableReadException(exception)) {
                        _diagnostics.Add(new EmailDiagnostic(
                            "OAB_ENTRY_SKIPPED",
                            exception.Message,
                            exception is OfflineAddressBookLimitExceededException
                                ? EmailDiagnosticSeverity.Warning
                                : EmailDiagnosticSeverity.Error,
                            reference.Id));
                        continue;
                    }
                    returned++;
                    yield return new OfflineAddressBookEntry(
                        reference, selected.Info, record.Properties, record.Diagnostics);
                }
            }
        }
    }

    private IEnumerable<OabAddressListSource> SelectSources(string? addressListId) {
        if (addressListId == null) return _sources;
        OabAddressListSource? selected = _sources.FirstOrDefault(source =>
            string.Equals(source.Info.Id, addressListId, StringComparison.Ordinal));
        if (selected == null) throw new KeyNotFoundException(string.Concat("OAB address list not found: ", addressListId));
        return new[] { selected };
    }

    private OabAddressListSource ValidateReference(OfflineAddressBookEntryReference reference) {
        if (reference.AddressListIndex < 0 || reference.AddressListIndex >= _sources.Count) {
            throw new ArgumentException("OAB entry reference belongs to another session snapshot.", nameof(reference));
        }
        OabAddressListSource selected = _sources[reference.AddressListIndex];
        if (selected.SnapshotId != reference.SnapshotId ||
            !string.Equals(selected.Info.Id, reference.AddressListId, StringComparison.Ordinal) ||
            reference.EntryIndex < 0 || reference.EntryIndex >= selected.Info.DeclaredEntryCount ||
            reference.RecordOffset < selected.Info.EntriesOffset || reference.RecordLength < 5 ||
            reference.RecordOffset > selected.Source.Length - reference.RecordLength) {
            throw new ArgumentException("OAB entry reference is invalid for this session snapshot.", nameof(reference));
        }
        return selected;
    }

    private void AddFramingDiagnostic(Exception exception, string location) {
        _diagnostics.Add(new EmailDiagnostic(
            "OAB_ENTRY_FRAMING_STOPPED",
            string.Concat(exception.Message, " Remaining records in this address list could not be located safely."),
            exception is OfflineAddressBookLimitExceededException
                ? EmailDiagnosticSeverity.Warning
                : EmailDiagnosticSeverity.Error,
            location));
    }

    private static bool IsRecoverableReadException(Exception exception) =>
        exception is InvalidDataException || exception is NotSupportedException ||
        exception is OfflineAddressBookLimitExceededException;

    private static string BuildEntryLocation(OfflineAddressBookListInfo info, long entryIndex) =>
        string.Concat(info.Id, ":", entryIndex.ToString("D10", CultureInfo.InvariantCulture));

    private void ThrowIfDisposed() {
        if (_disposed) throw new ObjectDisposedException(nameof(OfflineAddressBookSession));
    }
}
