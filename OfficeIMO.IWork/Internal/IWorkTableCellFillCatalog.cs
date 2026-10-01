namespace OfficeIMO.IWork.Internal;

/// <summary>Resolves only cell-style keys selected by validated modern cell storage.</summary>
internal sealed class IWorkTableCellFillCatalog(IWorkSourceDocument source, IWorkWireMessage store,
    IWorkArchiveRecord model, IWorkProjectionBudget budget, IWorkSourceReferenceIssueCollector references) {
    private readonly Dictionary<uint, (IWorkWireMessage Message, int Position)> _entries = new();
    private readonly Dictionary<uint, IWorkCellFill?> _resolved = new();
    private IWorkArchiveRecord? _list;
    private bool _initialized;

    internal bool FullyReconstructed { get; private set; } = true;

    internal IWorkCellFill? Read(uint key) {
        source.CancellationToken.ThrowIfCancellationRequested();
        if (!_initialized) Initialize();
        if (_resolved.TryGetValue(key, out var cached)) return cached;
        if (!_entries.TryGetValue(key, out var entry)) {
            FullyReconstructed = false;
            return null;
        }
        string path = IWorkTableCatalogIndex.EntryPath(entry.Position) + "/4";
        IWorkArchiveRecord? style = references.ReadOne(_list!, entry.Message, 4, path);
        bool complete = entry.Message.TotalFieldCount == entry.Message.FieldCount(1)
                + entry.Message.FieldCount(2) + entry.Message.FieldCount(4)
            && entry.Message.FieldCount(2) <= 1
            && !entry.Message.HasUnexpectedWireKind(2, IWorkWireKind.Varint)
            && entry.Message.FieldCount(4) == 1
            && !entry.Message.HasUnexpectedWireKind(4, IWorkWireKind.Bytes)
            && style?.MessageType == 6004;
        IWorkCellFill? fill = null;
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
                if (properties == null || !properties.HasField(1)) continue;
                IWorkWireMessage? declaration = IWorkObjectIndex.TryGetMessage(properties, 1, out bool malformedFill);
                bool valid = !malformedFill && properties.FieldCount(1) == 1
                    && !properties.HasUnexpectedWireKind(1, IWorkWireKind.Bytes) && declaration != null;
                IWorkColor? color = null;
                if (valid && declaration!.TotalFieldCount > 0) {
                    valid = declaration.TotalFieldCount == 1 && declaration.FieldCount(1) == 1
                        && IWorkColorReader.TryRead(declaration, 1, out color, ref valid)
                        && color is { Alpha: byte.MaxValue } && IsSupportedColor(declaration);
                }
                if (!valid) {
                    references.Declarations.Record(owner, "11/1", properties.FieldCount(1),
                        IWorkSourceDeclarationIssueKind.InvalidValue);
                    complete = false;
                    continue;
                }
                // An empty FillArchive clears the parent; it does not inherit the parent's color.
                fill = new IWorkCellFill(color);
            }
        }
        if (!complete) {
            references.Declarations.Record(_list!, path, entry.Message.FieldCount(4),
                IWorkSourceDeclarationIssueKind.RejectedMessageSet);
            FullyReconstructed = false;
            fill = null;
        }
        _resolved.Add(key, fill);
        return fill;
    }

    private static bool IsSupportedColor(IWorkWireMessage fill) {
        IWorkWireMessage color = IWorkObjectIndex.TryGetMessage(fill, 1)!;
        int[] fields = { 1, 3, 4, 5, 6, 11, 12 };
        if (color.TotalFieldCount != fields.Sum(color.FieldCount)) return false;
        foreach (int field in new[] { 1, 12 })
            if (color.FieldCount(field) > 1 || color.HasUnexpectedWireKind(field, IWorkWireKind.Varint)) return false;
        // Display P3 and CMYK are not silently reinterpreted as sRGB.
        if (color.HasField(12) && color.GetUnsigned(12) != 1) return false;
        return !color.HasField(1) || color.GetUnsigned(1) == (color.HasField(11) ? 3UL : 1UL);
    }

    private void Initialize() {
        _initialized = true;
        _list = references.ReadOne(model, store, 5, "4/5");
        if (store.FieldCount(5) != 1 || store.HasUnexpectedWireKind(5, IWorkWireKind.Bytes)
            || _list?.MessageType != 6005) {
            FullyReconstructed = false;
            return;
        }
        var declarations = IWorkTableCatalogIndex.Read(source, _list, budget, references, "cell-style");
        FullyReconstructed = declarations.IsComplete;
        if (!declarations.EnvelopeIsComplete) return;
        IWorkWireMessage message = source.Index.Message(_list);
        if (message.FieldCount(1) != 1 || message.HasUnexpectedWireKind(1, IWorkWireKind.Varint)
            || message.GetUnsigned(1) != 4) {
            references.Declarations.Record(_list, "1", message.FieldCount(1),
                IWorkSourceDeclarationIssueKind.InvalidSelectionMetadata);
            FullyReconstructed = false;
            return;
        }
        foreach (var entry in declarations.Entries) {
            source.CancellationToken.ThrowIfCancellationRequested();
            if (declarations.CanResolveKey(entry.Key)) _entries.Add(entry.Key, (entry.Message, entry.Position));
        }
    }
}
