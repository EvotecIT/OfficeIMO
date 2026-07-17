using OfficeIMO.Email;

namespace OfficeIMO.Email.Store;

internal sealed partial class PstStoreWriterCore : IDisposable {
    private const uint RootFolderNid = 0x122;
    private const uint IpmSubtreeNid = 0x8022;
    private const uint SearchRootNid = 0x8042;
    private const uint DeletedItemsNid = 0x8062;
    private const uint SpamSearchFolderNid = 0x2223;

    private readonly string _destinationPath;
    private readonly string _temporaryPath;
    private readonly EmailStorePstWriterOptions _options;
    private readonly PstWriterFile _file;
    private readonly Guid _providerUid;
    private readonly PstNamedPropertyWriter _namedProperties;
    private readonly PstWriterNodeJournal _nodes;
    private readonly PstWriterItemJournal _items;
    private readonly Dictionary<uint, FolderState> _folders = new Dictionary<uint, FolderState>();
    private readonly List<EmailStoreDiagnostic> _diagnostics = new List<EmailStoreDiagnostic>();
    private uint _nextFolderIndex = 0x10000;
    private uint _nextMessageIndex = 0x200000;
    private int _userFolderCount;
    private int _itemCount;
    private int _lastCheckpointItemCount;
    private bool _diagnosticsTruncated;
    private bool _completed;
    private bool _finalizing;
    private bool _abandon;
    private bool _disposed;

    internal PstStoreWriterCore(string destinationPath, EmailStorePstWriterOptions options) {
        _options = options ?? throw new ArgumentNullException(nameof(options));
        _providerUid = Guid.NewGuid();
        _namedProperties = new PstNamedPropertyWriter();
        _destinationPath = Path.GetFullPath(destinationPath);
        if (options.CheckpointPath != null && string.Equals(
            Path.GetFullPath(options.CheckpointPath), _destinationPath,
            StringComparison.OrdinalIgnoreCase)) {
            throw new InvalidOperationException(
                "The PST checkpoint and destination must use different paths.");
        }
        string? directory = Path.GetDirectoryName(_destinationPath);
        if (string.IsNullOrEmpty(directory)) directory = Directory.GetCurrentDirectory();
        Directory.CreateDirectory(directory);
        if (File.Exists(_destinationPath) && !options.OverwriteExisting) {
            throw new IOException("The destination PST already exists. Enable overwrite to replace it.");
        }
        _temporaryPath = Path.Combine(directory,
            string.Concat(".", Path.GetFileName(_destinationPath), ".",
                Guid.NewGuid().ToString("N"), ".tmp"));
        _nodes = new PstWriterNodeJournal(string.Concat(_temporaryPath, ".nodes"));
        _items = new PstWriterItemJournal(_temporaryPath);
        _file = new PstWriterFile(_temporaryPath);
        AddSystemFolder(RootFolderNid, RootFolderNid, "Root - Mailbox", null, false,
            EmailStoreSpecialFolderKind.Root);
        AddSystemFolder(IpmSubtreeNid, RootFolderNid, "Top of Personal Folders", "IPF.Note", false,
            EmailStoreSpecialFolderKind.IpmSubtree);
        AddSystemFolder(SearchRootNid, RootFolderNid, "Finder", null, false,
            EmailStoreSpecialFolderKind.SearchRoot);
        AddSystemFolder(DeletedItemsNid, IpmSubtreeNid, "Deleted Items", "IPF.Note", false,
            EmailStoreSpecialFolderKind.DeletedItems);
        AddSystemFolder(SpamSearchFolderNid, RootFolderNid, "SPAM Search Folder 2", "IPF.Note", true,
            EmailStoreSpecialFolderKind.Unknown);
        ReportProgress(EmailStorePstWriteStage.Initializing);
    }

    private PstStoreWriterCore(WriterCheckpointState state,
        EmailStorePstWriterOptions options) {
        _options = options;
        _destinationPath = state.DestinationPath;
        _temporaryPath = state.TemporaryPath;
        _providerUid = state.ProviderUid;
        _namedProperties = state.NamedProperties;
        _nextFolderIndex = state.NextFolderIndex;
        _nextMessageIndex = state.NextMessageIndex;
        _userFolderCount = state.UserFolderCount;
        _itemCount = state.ItemCount;
        _lastCheckpointItemCount = state.ItemCount;
        _diagnosticsTruncated = state.DiagnosticsTruncated;
        foreach (FolderState folder in state.Folders) _folders.Add(folder.Nid, folder);
        _diagnostics.AddRange(state.Diagnostics);
        _nodes = new PstWriterNodeJournal(string.Concat(_temporaryPath, ".nodes"),
            resume: true, state.NodeCount);
        _items = new PstWriterItemJournal(_temporaryPath, resume: true,
            state.ItemJournalCount, state.ItemPayloadLength);
        _file = new PstWriterFile(_temporaryPath, state.File);
        ReportProgress(EmailStorePstWriteStage.Initializing);
    }

    internal string RootFolderId => FormatId(IpmSubtreeNid);
    internal string MessageStoreRootFolderId => FormatId(RootFolderNid);
    internal string DeletedItemsFolderId => FormatId(DeletedItemsNid);
    internal string SearchRootFolderId => FormatId(SearchRootNid);
    internal string SpamSearchFolderId => FormatId(SpamSearchFolderNid);

    internal void SuppressWriterOwnedSpamSearchFolder() {
        ThrowIfUnavailable();
        _folders.Remove(SpamSearchFolderNid);
    }

    internal static bool IsWriterOwnedSearchFolderId(string id) => string.Equals(
        id, FormatId(SpamSearchFolderNid), StringComparison.Ordinal);

    internal string AddFolder(string name, string? parentFolderId, string? containerClass,
        EmailStoreSpecialFolderKind specialFolderKind) {
        ThrowIfUnavailable();
        if (string.IsNullOrWhiteSpace(name)) throw new ArgumentException("A folder name is required.", nameof(name));
        if (_userFolderCount >= _options.MaxFolderCount) {
            throw new InvalidOperationException("The configured PST folder limit has been reached.");
        }
        if (specialFolderKind != EmailStoreSpecialFolderKind.Unknown) {
            if (!CanAssignUserSpecialFolder(specialFolderKind)) {
                throw new ArgumentOutOfRangeException(nameof(specialFolderKind),
                    "The managed PST writer cannot assign this well-known role to a user folder.");
            }
            if (_folders.Values.Any(folder => folder.SpecialFolderKind == specialFolderKind)) {
                throw new InvalidOperationException(
                    "The PST writer already contains a folder with the requested well-known role.");
            }
        }
        uint parentNid = parentFolderId == null ? IpmSubtreeNid : ParseFolderId(parentFolderId);
        uint nid = AllocateNid(ref _nextFolderIndex, 0x02);
        _folders.Add(nid, new FolderState(nid, parentNid, name, containerClass, false,
            specialFolderKind));
        _userFolderCount++;
        ReportProgress(EmailStorePstWriteStage.WritingFolders);
        return FormatId(nid);
    }

    internal void ConfigureFolderMetadata(string folderId, string name, string? containerClass) {
        ThrowIfUnavailable();
        if (string.IsNullOrWhiteSpace(name)) {
            throw new ArgumentException("A folder name is required.", nameof(name));
        }
        FolderState folder = _folders[ParseFolderId(folderId)];
        folder.Name = name;
        folder.ContainerClass = containerClass;
    }

    internal string AddItem(string folderId, EmailDocument document, bool isAssociated,
        CancellationToken cancellationToken) {
        ThrowIfUnavailable();
        cancellationToken.ThrowIfCancellationRequested();
        if (_itemCount >= _options.MaxItemCount) {
            throw new InvalidOperationException("The configured PST item limit has been reached.");
        }
        uint folderNid = ParseFolderId(folderId);
        uint itemNid = AllocateNid(ref _nextMessageIndex, isAssociated ? 0x08U : 0x04U);
        WrittenMessage message = WriteMessage(document, itemNid, depth: 0, cancellationToken);
        _nodes.Add(new PstWriterNode(itemNid, folderNid, message.Context.DataBid,
            message.Context.SubnodeBid));
        IReadOnlyList<MapiProperty> tableProperties = SelectTableProperties(
            message.TableProperties, isAssociated ? AssociatedColumns : ContentsColumns);
        _items.Add(folderNid, itemNid, isAssociated, tableProperties);
        FolderState folder = _folders[folderNid];
        if (isAssociated) folder.AssociatedItemCount++;
        else {
            folder.NormalItemCount++;
            if (IsUnread(tableProperties)) folder.UnreadItemCount++;
        }
        _itemCount++;
        ReportProgress(EmailStorePstWriteStage.WritingItems);
        if (_options.CheckpointPath != null &&
            _itemCount - _lastCheckpointItemCount >= _options.CheckpointIntervalItems) {
            Checkpoint();
        }
        return FormatId(itemNid);
    }

    internal EmailStorePstWriteReport Complete(CancellationToken cancellationToken) {
        ThrowIfUnavailable();
        cancellationToken.ThrowIfCancellationRequested();
        _finalizing = true;
        ReportProgress(EmailStorePstWriteStage.Finalizing);
        using (PstWriterItemJournal.PstWriterItemSortedReader items =
            _items.OpenSorted(_options.MaxIndexRecordsInMemory)) {
            WriteStoreStructure(cancellationToken, items);
            if (!items.IsExhausted) {
                throw new InvalidDataException("The PST item spool contained an unmapped folder row.");
            }
        }
        if (_options.FailOnDataLoss && _diagnostics.Any(item =>
            item.Severity != EmailStoreDiagnosticSeverity.Information)) {
            throw new InvalidOperationException(
                "PST creation produced fidelity diagnostics and FailOnDataLoss is enabled.");
        }
        PstWriterTreeRoot nbt = _file.WriteNodeTree(
            _nodes.ReadSorted(_options.MaxIndexRecordsInMemory), _nodes.Count);
        PstWriterTreeRoot bbt = _file.WriteBlockTree();
        _file.FinalizeFile(nbt, bbt, _nodes.MaximumIndexes);
        _file.Dispose();
        _nodes.Dispose();
        _items.Dispose();
        CommitTemporaryFile();
        _completed = true;
        DeleteCheckpointFile();
        long bytes = new FileInfo(_destinationPath).Length;
        ReportProgress(EmailStorePstWriteStage.Completed);
        return new EmailStorePstWriteReport(_destinationPath, _userFolderCount,
            _itemCount, bytes, _diagnostics.ToArray(), _diagnosticsTruncated);
    }

    public void Dispose() {
        if (_disposed) return;
        _disposed = true;
        bool preserve = false;
        if (!_completed && !_abandon && _options.CheckpointPath != null &&
            _options.RetainCheckpointOnDispose) {
            if (!_finalizing) {
                try { CheckpointCore(); }
                catch (IOException) { }
                catch (UnauthorizedAccessException) { }
            }
            preserve = File.Exists(_options.CheckpointPath);
        }
        if (preserve) {
            _file.PreserveOnDispose();
            _nodes.PreserveOnDispose();
            _items.PreserveOnDispose();
        }
        _file.Dispose();
        _nodes.Dispose();
        _items.Dispose();
        if (!_completed && !preserve) {
            DeleteCheckpointFile();
            CleanupWorkingFiles(_temporaryPath, _destinationPath);
        }
    }

    internal static bool SupportsSpecialFolderKind(EmailStoreSpecialFolderKind kind) =>
        kind == EmailStoreSpecialFolderKind.Root ||
        kind == EmailStoreSpecialFolderKind.IpmSubtree ||
        kind == EmailStoreSpecialFolderKind.Inbox ||
        kind == EmailStoreSpecialFolderKind.Outbox ||
        kind == EmailStoreSpecialFolderKind.SentItems ||
        kind == EmailStoreSpecialFolderKind.DeletedItems ||
        kind == EmailStoreSpecialFolderKind.Drafts ||
        kind == EmailStoreSpecialFolderKind.Calendar ||
        kind == EmailStoreSpecialFolderKind.Contacts ||
        kind == EmailStoreSpecialFolderKind.Tasks ||
        kind == EmailStoreSpecialFolderKind.Notes ||
        kind == EmailStoreSpecialFolderKind.Journal ||
        kind == EmailStoreSpecialFolderKind.SearchRoot ||
        kind == EmailStoreSpecialFolderKind.CommonViews ||
        kind == EmailStoreSpecialFolderKind.PersonalViews;

    internal static bool CanAssignUserSpecialFolder(EmailStoreSpecialFolderKind kind) =>
        SupportsSpecialFolderKind(kind) &&
        kind != EmailStoreSpecialFolderKind.Root &&
        kind != EmailStoreSpecialFolderKind.IpmSubtree &&
        kind != EmailStoreSpecialFolderKind.DeletedItems &&
        kind != EmailStoreSpecialFolderKind.SearchRoot;

    private void AddSystemFolder(uint nid, uint parentNid, string name,
        string? containerClass, bool search, EmailStoreSpecialFolderKind specialFolderKind) =>
        _folders.Add(nid, new FolderState(nid, parentNid, name, containerClass, search,
            specialFolderKind));

    private uint ParseFolderId(string value) {
        if (value == null || value.Length != 12 ||
            !value.StartsWith("pst:", StringComparison.OrdinalIgnoreCase) ||
            !uint.TryParse(value.Substring(4), NumberStyles.HexNumber,
                CultureInfo.InvariantCulture, out uint nid) || !_folders.ContainsKey(nid)) {
            throw new ArgumentException("The folder identifier does not belong to this PST writer.", nameof(value));
        }
        return nid;
    }

    private void CommitTemporaryFile() {
        if (File.Exists(_destinationPath)) {
            if (!_options.OverwriteExisting) throw new IOException("The destination PST already exists.");
            File.Replace(_temporaryPath, _destinationPath, null);
        } else {
            File.Move(_temporaryPath, _destinationPath);
        }
    }

    private void ThrowIfUnavailable() {
        if (_disposed) throw new ObjectDisposedException(nameof(PstStoreWriterCore));
        if (_completed) throw new InvalidOperationException("The PST writer has already completed.");
    }

    private void Report(EmailStoreDiagnostic diagnostic) {
        if (_diagnostics.Count < _options.MaxDiagnostics) _diagnostics.Add(diagnostic);
        else _diagnosticsTruncated = true;
    }

    private static uint AllocateNid(ref uint nextIndex, uint type) {
        uint value = checked(nextIndex | type);
        nextIndex = checked(nextIndex + 0x20);
        return value;
    }

    private static string FormatId(uint nid) =>
        string.Concat("pst:", nid.ToString("X8", CultureInfo.InvariantCulture));

    private static void TryDelete(string path) {
        try { if (File.Exists(path)) File.Delete(path); }
        catch (IOException) { }
        catch (UnauthorizedAccessException) { }
    }

    private sealed class FolderState {
        internal FolderState(uint nid, uint parentNid, string name,
            string? containerClass, bool isSearchFolder,
            EmailStoreSpecialFolderKind specialFolderKind) {
            Nid = nid;
            ParentNid = parentNid;
            Name = name;
            ContainerClass = containerClass;
            IsSearchFolder = isSearchFolder;
            SpecialFolderKind = specialFolderKind;
        }
        internal uint Nid { get; }
        internal uint ParentNid { get; }
        internal string Name { get; set; }
        internal string? ContainerClass { get; set; }
        internal bool IsSearchFolder { get; }
        internal EmailStoreSpecialFolderKind SpecialFolderKind { get; }
        internal int NormalItemCount { get; set; }
        internal int AssociatedItemCount { get; set; }
        internal int UnreadItemCount { get; set; }
    }

    private readonly struct WrittenMessage {
        internal WrittenMessage(PstWriterContextResult context,
            IReadOnlyList<MapiProperty> tableProperties) {
            Context = context;
            TableProperties = tableProperties;
        }
        internal PstWriterContextResult Context { get; }
        internal IReadOnlyList<MapiProperty> TableProperties { get; }
    }
}
