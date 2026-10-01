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

    internal IWorkTableCellStyle? Read(uint key) {
        source.CancellationToken.ThrowIfCancellationRequested();
        if (!_initialized) Initialize();
        if (!_catalogComplete) FillsFullyReconstructed = LayoutFullyReconstructed = false;
        if (_resolved.TryGetValue(key, out var cached)) return cached;
        if (!_entries.TryGetValue(key, out var entry)) {
            FillsFullyReconstructed = LayoutFullyReconstructed = false;
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
        IWorkCellPadding? padding = null;
        IWorkCellVerticalAlignment? vertical = null;
        bool fillsComplete = true, paddingComplete = true, verticalComplete = true;
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
                ReadFill(owner, properties, ref fill, ref fillsComplete);
                ReadPadding(owner, properties, ref padding, ref paddingComplete);
                ReadVerticalAlignment(owner, properties, ref vertical, ref verticalComplete);
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
        var result = new IWorkTableCellStyle(fill, padding, vertical);
        _resolved.Add(key, result);
        return result;
    }

    private void ReadFill(IWorkArchiveRecord owner, IWorkWireMessage properties,
        ref IWorkCellFill? fill, ref bool complete) {
        if (!properties.HasField(1)) return;
        IWorkWireMessage? declaration = IWorkObjectIndex.TryGetMessage(properties, 1, out bool malformed);
        bool valid = !malformed && properties.FieldCount(1) == 1
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
            return;
        }
        // An empty FillArchive clears the parent rather than inheriting its color.
        fill = new IWorkCellFill(color);
    }

    private void ReadPadding(IWorkArchiveRecord owner, IWorkWireMessage properties,
        ref IWorkCellPadding? padding, ref bool complete) {
        if (!properties.HasField(9)) return;
        IWorkWireMessage? declaration = IWorkObjectIndex.TryGetMessage(properties, 9, out bool malformed);
        bool valid = !malformed && properties.FieldCount(9) == 1
            && !properties.HasUnexpectedWireKind(9, IWorkWireKind.Bytes) && declaration != null;
        var sides = new double[4];
        if (valid) {
            valid = declaration!.TotalFieldCount == Enumerable.Range(1, 4).Sum(declaration.FieldCount);
            for (int field = 1; field <= 4; field++) {
                float? value = declaration.GetFloat(field);
                if (declaration.FieldCount(field) > 1
                    || declaration.HasUnexpectedWireKind(field, IWorkWireKind.Fixed32)
                    || declaration.HasField(field) && !value.HasValue
                    || value.HasValue && (float.IsNaN(value.Value) || float.IsInfinity(value.Value) || value.Value < 0))
                    valid = false;
                sides[field - 1] = value ?? 0;
            }
        }
        if (!valid) {
            references.Declarations.Record(owner, "11/9", properties.FieldCount(9),
                IWorkSourceDeclarationIssueKind.InvalidValue);
            complete = false;
            return;
        }
        // The message overrides the whole property. Omitted scalar sides use protobuf zero defaults.
        padding = new IWorkCellPadding(sides[0], sides[1], sides[2], sides[3]);
    }

    private void ReadVerticalAlignment(IWorkArchiveRecord owner, IWorkWireMessage properties,
        ref IWorkCellVerticalAlignment? vertical, ref bool complete) {
        if (!properties.HasField(8)) return;
        ulong? value = properties.GetUnsigned(8);
        if (properties.FieldCount(8) != 1 || properties.HasUnexpectedWireKind(8, IWorkWireKind.Varint)
            || !value.HasValue || value.Value > 2) {
            references.Declarations.Record(owner, "11/8", properties.FieldCount(8),
                IWorkSourceDeclarationIssueKind.InvalidValue);
            complete = false;
            return;
        }
        vertical = (IWorkCellVerticalAlignment)value.Value;
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

    /// <summary>Resolves a selected text-style entry without traversing unselected style references.</summary>
    internal IWorkArchiveRecord? ReadTextStyle(uint key, ref bool complete) {
        source.CancellationToken.ThrowIfCancellationRequested();
        if (!_initialized) Initialize();
        if (!_catalogComplete) complete = false;
        if (!_entries.TryGetValue(key, out var entry)) { complete = false; return null; }
        string path = IWorkTableCatalogIndex.EntryPath(entry.Position) + "/4";
        IWorkArchiveRecord? record = references.ReadOne(_list!, entry.Message, 4, path);
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
        _list = references.ReadOne(model, store, 5, "4/5");
        if (store.FieldCount(5) != 1 || store.HasUnexpectedWireKind(5, IWorkWireKind.Bytes)
            || _list?.MessageType != 6005) {
            _catalogComplete = false;
            return;
        }
        var declarations = IWorkTableCatalogIndex.Read(source, _list, budget, references, "cell-style");
        _catalogComplete = declarations.IsComplete;
        if (!declarations.EnvelopeIsComplete) return;
        IWorkWireMessage message = source.Index.Message(_list);
        if (message.FieldCount(1) != 1 || message.HasUnexpectedWireKind(1, IWorkWireKind.Varint)
            || message.GetUnsigned(1) != 4) {
            references.Declarations.Record(_list, "1", message.FieldCount(1),
                IWorkSourceDeclarationIssueKind.InvalidSelectionMetadata);
            _catalogComplete = false;
            return;
        }
        foreach (var entry in declarations.Entries) {
            source.CancellationToken.ThrowIfCancellationRequested();
            if (declarations.CanResolveKey(entry.Key)) _entries.Add(entry.Key, (entry.Message, entry.Position));
        }
    }
}

/// <summary>Resolved independent selected cell properties; unsupported properties stay null without discarding other qualified fields.</summary>
internal sealed class IWorkTableCellStyle(IWorkCellFill? fill, IWorkCellPadding? padding,
    IWorkCellVerticalAlignment? verticalAlignment) {
    internal IWorkCellFill? Fill { get; } = fill;
    internal IWorkCellPadding? Padding { get; } = padding;
    internal IWorkCellVerticalAlignment? VerticalAlignment { get; } = verticalAlignment;
}
