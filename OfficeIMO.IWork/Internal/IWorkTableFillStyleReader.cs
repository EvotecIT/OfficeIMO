namespace OfficeIMO.IWork.Internal;

/// <summary>Reads applicable unbanded role fills through the shared inheritance and fill decoders.</summary>
internal sealed class IWorkTableFillStyleReader {
    internal bool FullyReconstructed { get; private set; } = true;
    internal IWorkTableFillStyles Defaults { get; }

    internal IWorkTableFillStyleReader(IWorkSourceDocument source, IWorkArchiveRecord model,
        IWorkWireMessage message, IWorkProjectionBudget budget, IWorkSourceReferenceIssueCollector references,
        int rows, int columns, int headerRows, int headerColumns, int footerRows, bool defaultsSupported) {
        bool hasCells = rows > 0 && columns > 0;
        if (!defaultsSupported) {
            FullyReconstructed = !hasCells;
            Defaults = new IWorkTableFillStyles(null, null, null, null);
            return;
        }
        var resolved = new Dictionary<ulong, (IWorkCellFill? Fill, bool Complete)>();
        Defaults = new IWorkTableFillStyles(
            ReadRole(18, hasCells && rows > headerRows + footerRows && columns > headerColumns),
            ReadRole(19, hasCells && headerRows > 0),
            ReadRole(20, hasCells && rows > headerRows && headerColumns > 0),
            ReadRole(21, hasCells && footerRows > 0 && columns > headerColumns));

        IWorkCellFill? ReadRole(int field, bool applicable) {
            if (!applicable || !message.HasField(field)) return null;
            source.CancellationToken.ThrowIfCancellationRequested();
            IWorkArchiveRecord? record = references.ReadOne(model, message, field);
            bool complete = message.FieldCount(field) == 1
                && !message.HasUnexpectedWireKind(field, IWorkWireKind.Bytes) && record?.MessageType == 6004;
            IWorkCellFill? fill = null;
            if (complete) {
                if (resolved.TryGetValue(record!.Identifier, out var cached)) {
                    fill = cached.Fill; complete = cached.Complete;
                } else {
                    var chain = IWorkStyleReader.ReadChain(source.Index, record!.Identifier,
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
                            complete = false; continue;
                        }
                        if (properties != null) IWorkCellFillReader.Read(owner, properties, references, ref fill, ref complete);
                    }
                    if (!complete) fill = null;
                    resolved.Add(record.Identifier, (fill, complete));
                }
            }
            if (!complete) {
                FullyReconstructed = false;
                references.Declarations.Record(model, field.ToString(System.Globalization.CultureInfo.InvariantCulture),
                    message.FieldCount(field), IWorkSourceDeclarationIssueKind.RejectedMessageSet);
            }
            return fill;
        }
    }
}
