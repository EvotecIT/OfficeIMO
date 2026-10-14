namespace OfficeIMO.IWork.Internal;

/// <summary>Reads applicable region layouts without allocating empty cells or traversing unused roles.</summary>
internal sealed class IWorkTableCellLayoutReader {
    internal bool FullyReconstructed { get; private set; } = true;
    internal IReadOnlyList<IWorkCellLayout?> Defaults { get; }

    internal IWorkTableCellLayoutReader(IWorkSourceDocument source, IWorkArchiveRecord model,
        IWorkWireMessage message, IWorkProjectionBudget budget, IWorkSourceReferenceIssueCollector references,
        int rows, int columns, int headerRows, int headerColumns, int footerRows) {
        bool hasCells = rows > 0 && columns > 0;
        var resolved = new Dictionary<ulong, (IWorkCellLayout Layout, bool Complete)>();
        Defaults = Array.AsReadOnly(new[] {
            ReadRole(18, hasCells && rows > headerRows + footerRows && columns > headerColumns),
            ReadRole(19, hasCells && headerRows > 0),
            ReadRole(20, hasCells && rows > headerRows + footerRows && headerColumns > 0),
            ReadRole(21, hasCells && footerRows > 0)
        });

        IWorkCellLayout? ReadRole(int field, bool applicable) {
            if (!applicable || !message.HasField(field)) return null;
            source.CancellationToken.ThrowIfCancellationRequested();
            IWorkArchiveRecord? record = references.ReadOne(model, message, field, allowedType: static type => type == 6004);
            bool complete = message.FieldCount(field) == 1
                && !message.HasUnexpectedWireKind(field, IWorkWireKind.Bytes) && record?.MessageType == 6004;
            IWorkCellLayout? layout = null;
            if (complete) {
                if (resolved.TryGetValue(record!.Identifier, out var cached)) {
                    layout = cached.Layout; complete = cached.Complete;
                } else {
                    var chain = IWorkStyleReader.ReadChain(source.Index, record!.Identifier,
                        budget.MaximumTextStyleInheritanceDepth, type => type == 6004,
                        tolerateStyleDepth: true, references, ref complete);
                    IWorkCellPadding? padding = null;
                    IWorkCellVerticalAlignment? vertical = null;
                    bool paddingComplete = true, verticalComplete = true;
                    var reader = new IWorkCellLayoutReader(references);
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
                        reader.ReadPadding(owner, properties, ref padding, ref paddingComplete);
                        reader.ReadVerticalAlignment(owner, properties, ref vertical, ref verticalComplete);
                    }
                    layout = new IWorkCellLayout(complete && paddingComplete ? padding : null,
                        complete && verticalComplete ? vertical : null);
                    complete &= paddingComplete && verticalComplete;
                    resolved.Add(record.Identifier, (layout, complete));
                }
            }
            if (!complete) {
                FullyReconstructed = false;
                references.Declarations.Record(model, field.ToString(System.Globalization.CultureInfo.InvariantCulture),
                    message.FieldCount(field), IWorkSourceDeclarationIssueKind.RejectedMessageSet);
            }
            return layout;
        }
    }
}
