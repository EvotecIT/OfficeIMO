using OfficeIMO.Email;
using System.Security.Cryptography;

namespace OfficeIMO.Email.Store;

internal sealed class EmailStorePstMerger {
    private readonly EmailStoreMergeSource[] _sources;
    private readonly string _destination;
    private readonly EmailStorePstMergeOptions _options;
    private readonly CancellationToken _cancellationToken;
    private readonly List<EmailStoreDiagnostic> _diagnostics = new List<EmailStoreDiagnostic>();
    private readonly List<EmailStoreMergeSourceReport> _sourceReports = new List<EmailStoreMergeSourceReport>();
    private readonly Dictionary<string, DestinationFolder> _mergedFolders =
        new Dictionary<string, DestinationFolder>(StringComparer.OrdinalIgnoreCase);
    private readonly HashSet<string> _sourceRootNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
    private bool _diagnosticsTruncated;
    private int _inspected;
    private int _written;
    private int _duplicates;
    private int _skipped;
    private int _retries;

    internal EmailStorePstMerger(IEnumerable<EmailStoreMergeSource> sources, string destinationPath,
        EmailStorePstMergeOptions options, CancellationToken cancellationToken) {
        if (sources == null) throw new ArgumentNullException(nameof(sources));
        _sources = sources.ToArray();
        if (_sources.Length == 0) throw new ArgumentException("At least one merge source is required.", nameof(sources));
        if (_sources.Any(source => source == null)) throw new ArgumentException("A merge source cannot be null.", nameof(sources));
        _destination = Path.GetFullPath(destinationPath);
        _options = options;
        _cancellationToken = cancellationToken;
        foreach (EmailStoreMergeSource source in _sources) {
            if (EmailStorePathIdentity.AreEquivalent(source.Path, _destination)) {
                throw new InvalidOperationException("A merge source cannot also be the destination PST.");
            }
            if (Directory.Exists(source.Path) &&
                EmailStorePathIdentity.IsSameOrDescendant(_destination, source.Path)) {
                throw new InvalidOperationException(
                    "The destination PST cannot be created inside a mailbox-directory source.");
            }
        }
    }

    internal EmailStorePstMergeReport Run() {
        var writerOptions = new EmailStorePstWriterOptions(
            _options.DisplayName,
            _options.OverwriteExisting,
            failOnDataLoss: false,
            _options.MaxFolderCount,
            _options.MaxItems,
            _options.MaxNestedMessageDepth,
            maxIndexRecordsInMemory: _options.MaxIndexRecordsInMemory,
            retainCheckpointOnDispose: false,
            maxDiagnostics: _options.MaxDiagnostics);
        using var writer = EmailStorePstWriter.Create(_destination, writerOptions);
        using EmailSemanticDedupIndex? deduplication = _options.Deduplicate
            ? new EmailSemanticDedupIndex(_destination)
            : null;

        for (int sourceIndex = 0; sourceIndex < _sources.Length; sourceIndex++) {
            _cancellationToken.ThrowIfCancellationRequested();
            ProcessSource(writer, deduplication, sourceIndex);
            if (_inspected >= _options.MaxItems) break;
        }

        ReportProgress(EmailStoreMergeStage.Finalizing, Math.Max(0, _sources.Length - 1));
        EmailStorePstWriteReport writeReport = writer.Complete(_cancellationToken);
        foreach (EmailStoreDiagnostic diagnostic in writeReport.Diagnostics) AddDiagnostic(diagnostic);
        ReportProgress(EmailStoreMergeStage.Completed, Math.Max(0, _sources.Length - 1));
        return new EmailStorePstMergeReport(writeReport, _sourceReports.AsReadOnly(),
            _inspected, _written, _duplicates, _skipped, _retries,
            _diagnostics.AsReadOnly(), _diagnosticsTruncated || writeReport.DiagnosticsTruncated);
    }

    private void ProcessSource(EmailStorePstWriter writer, EmailSemanticDedupIndex? deduplication,
        int sourceIndex) {
        EmailStoreMergeSource source = _sources[sourceIndex];
        int inspectedBefore = _inspected;
        int writtenBefore = _written;
        int duplicateBefore = _duplicates;
        int skippedBefore = _skipped;
        int retriesBefore = _retries;
        EmailStoreFormat format = EmailStoreFormat.Unknown;
        int folderCount = 0;
        bool completed = false;
        bool destinationWriteInProgress = false;
        ReportProgress(EmailStoreMergeStage.OpeningSource, sourceIndex);
        try {
            using EmailStoreSession session = OpenWithRetry(source);
            format = session.Format;
            folderCount = session.Folders.Count;
            foreach (EmailStoreDiagnostic diagnostic in session.Diagnostics) AddDiagnostic(diagnostic);
            ReportProgress(EmailStoreMergeStage.MappingFolders, sourceIndex);
            Dictionary<string, string> folderMap = CreateFolderMap(writer, session, source, sourceIndex);
            ReportProgress(EmailStoreMergeStage.WritingItems, sourceIndex);
            int remaining = _options.MaxItems - _inspected;
            if (remaining <= 0) return;
            var enumeration = new EmailStoreEnumerationOptions(
                includeAssociatedItems: _options.IncludeAssociatedItems,
                includeOrphanedItems: _options.IncludeOrphanedItems,
                maxItems: remaining);
            var readOptions = new EmailStoreItemReadOptions(
                EmailStoreItemReadParts.All, preferStreamingAttachmentContent: true);
            foreach (EmailStoreItemReference reference in session.EnumerateItems(
                enumeration, _cancellationToken)) {
                _cancellationToken.ThrowIfCancellationRequested();
                _inspected++;
                if (!folderMap.TryGetValue(reference.FolderId, out string? destinationFolder)) {
                    _skipped++;
                    AddDiagnostic(new EmailStoreDiagnostic(
                        "EMAIL_STORE_MERGE_FOLDER_UNMAPPED",
                        "An item was skipped because its source folder was excluded or could not be mapped.",
                        EmailStoreDiagnosticSeverity.Warning));
                    continue;
                }
                PreparedItem prepared;
                try {
                    prepared = PrepareWithRetry(session, reference, readOptions,
                        deduplication != null);
                } catch (Exception exception) when (CanSkipItem(exception)) {
                    _skipped++;
                    AddDiagnostic(new EmailStoreDiagnostic(
                        "EMAIL_STORE_MERGE_ITEM_SKIPPED",
                        string.Concat("A source item could not be merged: ", exception.Message),
                        EmailStoreDiagnosticSeverity.Error));
                    if (!_options.ContinueOnItemError) throw;
                    ReportProgress(EmailStoreMergeStage.WritingItems, sourceIndex);
                    continue;
                }
                if (prepared.Digest != null && deduplication!.Contains(prepared.Digest)) {
                    _duplicates++;
                    ReportProgress(EmailStoreMergeStage.WritingItems, sourceIndex);
                    continue;
                }

                // Destination mutations are intentionally outside the skippable source-read boundary.
                // Any writer or dedup-index failure aborts the atomic merge instead of continuing from
                // an uncertain intermediate state.
                destinationWriteInProgress = true;
                writer.AddItem(destinationFolder, prepared.Item.Document, reference.IsAssociated,
                    _cancellationToken);
                if (prepared.Digest != null && !deduplication!.Add(prepared.Digest)) {
                    throw new InvalidDataException("The semantic deduplication index changed unexpectedly.");
                }
                destinationWriteInProgress = false;
                _written++;
                if (reference.IsOrphaned) {
                    AddDiagnostic(new EmailStoreDiagnostic(
                        "EMAIL_STORE_MERGE_ORPHAN_RECOVERED",
                        "An item absent from its source contents table was recovered and merged.",
                        EmailStoreDiagnosticSeverity.Information));
                }
                ReportProgress(EmailStoreMergeStage.WritingItems, sourceIndex);
                if (_inspected >= _options.MaxItems) break;
            }
            completed = true;
        } catch (Exception exception) when (!destinationWriteInProgress && CanSkipSource(exception)) {
            AddDiagnostic(new EmailStoreDiagnostic(
                "EMAIL_STORE_MERGE_SOURCE_SKIPPED",
                string.Concat("A source could not be merged: ", exception.Message),
                EmailStoreDiagnosticSeverity.Error));
            if (!_options.ContinueOnSourceError) throw;
        } finally {
            _sourceReports.Add(new EmailStoreMergeSourceReport(source.Path, format, folderCount,
                _inspected - inspectedBefore, _written - writtenBefore,
                _duplicates - duplicateBefore, _skipped - skippedBefore,
                _retries - retriesBefore, completed));
        }
    }

    private EmailStoreSession OpenWithRetry(EmailStoreMergeSource source) {
        int attempt = 0;
        while (true) {
            _cancellationToken.ThrowIfCancellationRequested();
            try {
                return EmailStoreSession.Open(source.Path, source.ReaderOptions, _cancellationToken);
            } catch (IOException) when (attempt < _options.MaxRetries) {
                attempt++;
                _retries++;
                DelayRetry();
            }
        }
    }

    private PreparedItem PrepareWithRetry(EmailStoreSession session,
        EmailStoreItemReference reference, EmailStoreItemReadOptions readOptions,
        bool createDigest) {
        int attempt = 0;
        while (true) {
            _cancellationToken.ThrowIfCancellationRequested();
            try {
                EmailStoreItem item = session.ReadItem(reference, readOptions, _cancellationToken);
                byte[]? digest = createDigest
                    ? CreateDeduplicationDigest(item.Document, reference.IsAssociated)
                    : null;
                return new PreparedItem(item, digest);
            } catch (IOException) when (attempt < _options.MaxRetries) {
                attempt++;
                _retries++;
                DelayRetry();
            }
        }
    }

    private Dictionary<string, string> CreateFolderMap(EmailStorePstWriter writer,
        EmailStoreSession session, EmailStoreMergeSource source, int sourceIndex) {
        var map = new Dictionary<string, string>(StringComparer.Ordinal);
        if (_options.FolderMode == EmailStoreMergeFolderMode.Flatten) {
            foreach (EmailStoreFolderInfo folder in session.Folders) {
                if (!folder.IsSearchFolder || _options.IncludeSearchFolders) map[folder.Id] = writer.RootFolderId;
            }
            return map;
        }

        string baseFolder = writer.RootFolderId;
        if (_options.FolderMode == EmailStoreMergeFolderMode.SeparateSourceRoots) {
            string label = CreateUniqueSourceRootName(source, session, sourceIndex);
            baseFolder = writer.AddFolder(label, writer.RootFolderId);
        }
        foreach (EmailStoreFolderInfo folder in session.Folders) {
            if (folder.SpecialFolderKind == EmailStoreSpecialFolderKind.Root ||
                folder.SpecialFolderKind == EmailStoreSpecialFolderKind.IpmSubtree) {
                map[folder.Id] = baseFolder;
            }
        }

        var pending = session.Folders.Where(folder => !map.ContainsKey(folder.Id)).ToList();
        bool madeProgress;
        do {
            madeProgress = false;
            for (int index = pending.Count - 1; index >= 0; index--) {
                EmailStoreFolderInfo folder = pending[index];
                if (folder.IsSearchFolder && !_options.IncludeSearchFolders) {
                    pending.RemoveAt(index);
                    madeProgress = true;
                    continue;
                }
                string parent;
                if (folder.ParentId == null) parent = baseFolder;
                else if (!map.TryGetValue(folder.ParentId, out parent!)) continue;
                map[folder.Id] = MapFolder(writer, parent, folder);
                if (folder.IsSearchFolder) {
                    AddDiagnostic(new EmailStoreDiagnostic(
                        "EMAIL_STORE_MERGE_SEARCH_FOLDER_STATIC",
                        "A source search folder was merged as a static folder.",
                        EmailStoreDiagnosticSeverity.Warning));
                }
                pending.RemoveAt(index);
                madeProgress = true;
            }
        } while (madeProgress && pending.Count > 0);

        foreach (EmailStoreFolderInfo folder in pending) {
            if (folder.IsSearchFolder && !_options.IncludeSearchFolders) continue;
            map[folder.Id] = MapFolder(writer, baseFolder, folder);
            AddDiagnostic(new EmailStoreDiagnostic(
                "EMAIL_STORE_MERGE_FOLDER_PARENT_RECOVERED",
                "A folder with an unavailable or cyclic parent was attached to its source root.",
                EmailStoreDiagnosticSeverity.Warning));
        }
        return map;
    }

    private string MapFolder(EmailStorePstWriter writer, string parent, EmailStoreFolderInfo folder) {
        if (_options.FolderMode != EmailStoreMergeFolderMode.MergeByFolderPath) {
            return writer.AddFolder(folder.Name, parent, folder.ContainerClass);
        }
        string key = string.Concat(parent, "\0", folder.Name);
        if (_mergedFolders.TryGetValue(key, out DestinationFolder? existing)) {
            if (!string.Equals(existing.ContainerClass, folder.ContainerClass,
                StringComparison.OrdinalIgnoreCase)) {
                AddDiagnostic(new EmailStoreDiagnostic(
                    "EMAIL_STORE_MERGE_FOLDER_CLASS_CONFLICT",
                    "Equivalent folder paths declared different container classes; the first declaration was retained.",
                    EmailStoreDiagnosticSeverity.Warning));
            }
            return existing.Id;
        }
        string id = writer.AddFolder(folder.Name, parent, folder.ContainerClass);
        _mergedFolders.Add(key, new DestinationFolder(id, folder.ContainerClass));
        return id;
    }

    private string CreateUniqueSourceRootName(EmailStoreMergeSource source,
        EmailStoreSession session, int sourceIndex) {
        string label = source.DisplayName ?? session.DisplayName ?? SourceFileName(source.Path);
        if (string.IsNullOrWhiteSpace(label)) {
            label = string.Concat("Source ", (sourceIndex + 1).ToString(CultureInfo.InvariantCulture));
        }
        string candidate = label.Trim();
        int suffix = 2;
        while (!_sourceRootNames.Add(candidate)) {
            candidate = string.Concat(label.Trim(), " [", suffix.ToString(CultureInfo.InvariantCulture), "]");
            suffix++;
        }
        return candidate;
    }

    private byte[] CreateDeduplicationDigest(EmailDocument document, bool associated) {
        byte[] semantic = EmailSemanticComparer.CreateFingerprint(
            document, _options.DeduplicationOptions, _cancellationToken).Digest;
        using (SHA256 sha = SHA256.Create()) {
            var domain = new byte[semantic.Length + 1];
            domain[0] = associated ? (byte)1 : (byte)0;
            Buffer.BlockCopy(semantic, 0, domain, 1, semantic.Length);
            return sha.ComputeHash(domain);
        }
    }

    private void DelayRetry() {
        if (_options.RetryDelay == TimeSpan.Zero) return;
        Task.Delay(_options.RetryDelay, _cancellationToken).GetAwaiter().GetResult();
    }

    private void ReportProgress(EmailStoreMergeStage stage, int sourceIndex) =>
        _options.Progress?.Report(new EmailStoreMergeProgress(stage, sourceIndex, _sources.Length,
            _inspected, _written, _duplicates, _skipped));

    private void AddDiagnostic(EmailStoreDiagnostic diagnostic) {
        if (_diagnostics.Count < _options.MaxDiagnostics) _diagnostics.Add(diagnostic);
        else _diagnosticsTruncated = true;
    }

    private bool CanSkipItem(Exception exception) =>
        _options.ContinueOnItemError && (exception is IOException || exception is InvalidDataException ||
            exception is NotSupportedException || exception is EmailStoreLimitExceededException);

    private bool CanSkipSource(Exception exception) =>
        _options.ContinueOnSourceError && (exception is IOException || exception is InvalidDataException ||
            exception is NotSupportedException || exception is UnauthorizedAccessException ||
            exception is EmailStoreLimitExceededException);

    private static string SourceFileName(string path) {
        string trimmed = path.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
        string name = Path.GetFileNameWithoutExtension(trimmed);
        return string.IsNullOrWhiteSpace(name) ? Path.GetFileName(trimmed) : name;
    }

    private sealed class DestinationFolder {
        internal DestinationFolder(string id, string? containerClass) {
            Id = id;
            ContainerClass = containerClass;
        }
        internal string Id { get; }
        internal string? ContainerClass { get; }
    }

    private readonly struct PreparedItem {
        internal PreparedItem(EmailStoreItem item, byte[]? digest) {
            Item = item;
            Digest = digest;
        }
        internal EmailStoreItem Item { get; }
        internal byte[]? Digest { get; }
    }
}
