namespace OfficeIMO.IWork.Internal;

/// <summary>Resolves only cell-style keys selected by validated modern cell storage.</summary>
internal sealed class IWorkTableCellStyleCatalog(IWorkSourceDocument source, IWorkWireMessage store,
    IWorkArchiveRecord model, IWorkProjectionBudget budget, IWorkSourceReferenceIssueCollector references) {
    private readonly Dictionary<uint, (IWorkWireMessage Message, int Position)> _entries = new();
    private readonly Dictionary<uint, IWorkTableCellStyle?> _resolved = new();
    private IWorkArchiveRecord? _list;
    private bool _initialized;
    private bool _catalogComplete = true;

    internal bool FillsFullyReconstructed { get; private set; } = true;
    internal bool LayoutFullyReconstructed { get; private set; } = true;
    internal bool BordersFullyReconstructed { get; private set; } = true;

    internal IWorkTableCellStyle? Read(uint key) {
        source.CancellationToken.ThrowIfCancellationRequested();
        if (!_initialized) Initialize();
        if (!_catalogComplete) FillsFullyReconstructed = LayoutFullyReconstructed = BordersFullyReconstructed = false;
        if (_resolved.TryGetValue(key, out var cached)) return cached;
        if (!_entries.TryGetValue(key, out var entry)) {
            FillsFullyReconstructed = LayoutFullyReconstructed = BordersFullyReconstructed = false;
            return null;
        }
        string path = IWorkTableCatalogIndex.EntryPath(entry.Position) + "/4";
        IWorkArchiveRecord? style = references.ReadOne(_list!, entry.Message, 4, path, static type => type == 6004);
        bool complete = entry.Message.TotalFieldCount == entry.Message.FieldCount(1)
                + entry.Message.FieldCount(2) + entry.Message.FieldCount(4)
            && entry.Message.FieldCount(2) <= 1
            && !entry.Message.HasUnexpectedWireKind(2, IWorkWireKind.Varint)
            && entry.Message.FieldCount(4) == 1
            && !entry.Message.HasUnexpectedWireKind(4, IWorkWireKind.Bytes)
            && style?.MessageType == 6004;
        IWorkCellFill? fill = null;
        IWorkCellPadding? padding = null;
        IWorkCellVerticalAlignment? vertical = null;
        bool fillsComplete = true, paddingComplete = true, verticalComplete = true;
        var borders = new IWorkCellBorderStyle();
        var layoutReader = new IWorkCellLayoutReader(references);
        if (complete) {
            var chain = IWorkStyleReader.ReadChain(source.Index, style!.Identifier,
                budget.MaximumTextStyleInheritanceDepth, type => type == 6004,
                tolerateStyleDepth: true, references, ref complete);
            for (int i = chain.Count - 1; i >= 0; i--) {
                source.CancellationToken.ThrowIfCancellationRequested();
                var (owner, payload) = chain[i];
                IWorkWireMessage? properties = IWorkObjectIndex.TryGetMessage(payload, 11, out bool malformed);
                if (malformed || payload.FieldCount(11) > 1
                    || payload.HasUnexpectedWireKind(11, IWorkWireKind.Bytes)
                    || payload.HasField(11) && properties == null) {
                    references.Declarations.Record(owner, "11", payload.FieldCount(11));
                    complete = false;
                    continue;
                }
                if (properties == null) continue;
                IWorkCellFillReader.Read(owner, properties, references, ref fill, ref fillsComplete);
                layoutReader.ReadPadding(owner, properties, ref padding, ref paddingComplete);
                layoutReader.ReadVerticalAlignment(owner, properties, ref vertical, ref verticalComplete);
                borders.Read(owner, properties, references);
            }
        }
        if (!complete || !fillsComplete || !paddingComplete || !verticalComplete) {
            references.Declarations.Record(_list!, path, entry.Message.FieldCount(4),
                IWorkSourceDeclarationIssueKind.RejectedMessageSet);
        }
        if (!complete || !fillsComplete) {
            FillsFullyReconstructed = false;
            fill = null;
        }
        if (!complete || !paddingComplete || !verticalComplete) LayoutFullyReconstructed = false;
        if (!complete || !paddingComplete) padding = null;
        if (!complete || !verticalComplete) vertical = null;
        if (!complete) borders.Invalidate();
        if (!borders.FullyReconstructed) BordersFullyReconstructed = false;
        var result = new IWorkTableCellStyle(fill, padding, vertical, complete && fillsComplete, borders,
            complete && paddingComplete, complete && verticalComplete);
        _resolved.Add(key, result);
        return result;
    }

    /// <summary>Resolves a selected text-style entry without traversing unselected style references.</summary>
    internal IWorkArchiveRecord? ReadTextStyle(uint key, ref bool complete) {
        source.CancellationToken.ThrowIfCancellationRequested();
        if (!_initialized) Initialize();
        if (!_catalogComplete) complete = false;
        if (!_entries.TryGetValue(key, out var entry)) { complete = false; return null; }
        string path = IWorkTableCatalogIndex.EntryPath(entry.Position) + "/4";
        IWorkArchiveRecord? record = references.ReadOne(_list!, entry.Message, 4, path, static type => type == 2022);
        if (entry.Message.TotalFieldCount != entry.Message.FieldCount(1)
                + entry.Message.FieldCount(2) + entry.Message.FieldCount(4)
            || entry.Message.FieldCount(2) > 1
            || entry.Message.HasUnexpectedWireKind(2, IWorkWireKind.Varint)
            || entry.Message.FieldCount(4) != 1
            || entry.Message.HasUnexpectedWireKind(4, IWorkWireKind.Bytes)
            || record?.MessageType != 2022) {
            references.Declarations.Record(_list!, path, entry.Message.FieldCount(4),
                IWorkSourceDeclarationIssueKind.RejectedMessageSet);
            complete = false;
            return null;
        }
        return record;
    }

    internal void RecordTextStyleFailure(uint key) {
        if (_list != null && _entries.TryGetValue(key, out var entry))
            references.Declarations.Record(_list, IWorkTableCatalogIndex.EntryPath(entry.Position) + "/4",
                entry.Message.FieldCount(4), IWorkSourceDeclarationIssueKind.RejectedMessageSet);
    }

    private void Initialize() {
        _initialized = true;
        _list = references.ReadOne(model, store, 5, "4/5", IWorkTableCatalogIndex.IsDataListType);
        if (store.FieldCount(5) != 1 || store.HasUnexpectedWireKind(5, IWorkWireKind.Bytes)
            || _list == null || !IWorkTableCatalogIndex.IsDataListType(_list.MessageType)) {
            _catalogComplete = false;
            return;
        }
        var declarations = IWorkTableCatalogIndex.Read(source, _list, budget, references, "cell-style", 4);
        _catalogComplete = declarations.IsComplete;
        if (!declarations.EnvelopeIsComplete) return;
        foreach (var entry in declarations.Entries) {
            source.CancellationToken.ThrowIfCancellationRequested();
            if (declarations.CanResolveKey(entry.Key)) _entries.Add(entry.Key, (entry.Message, entry.Position));
        }
    }
}

/// <summary>Resolved independent selected cell properties; unsupported properties stay null without discarding other qualified fields.</summary>
internal sealed class IWorkTableCellStyle(IWorkCellFill? fill, IWorkCellPadding? padding,
    IWorkCellVerticalAlignment? verticalAlignment, bool fillIsComplete = true, IWorkCellBorderStyle? borders = null,
    bool paddingIsComplete = true, bool verticalAlignmentIsComplete = true) {
    internal bool FillIsComplete { get; } = fillIsComplete;
    internal bool PaddingIsComplete { get; } = paddingIsComplete;
    internal bool VerticalAlignmentIsComplete { get; } = verticalAlignmentIsComplete;
    internal IWorkCellFill? Fill { get; } = fill;
    internal IWorkCellPadding? Padding { get; } = padding;
    internal IWorkCellVerticalAlignment? VerticalAlignment { get; } = verticalAlignment;
    internal IWorkCellBorderStyle? Borders { get; } = borders;
}
