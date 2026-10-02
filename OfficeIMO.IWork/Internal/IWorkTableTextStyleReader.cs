namespace OfficeIMO.IWork.Internal;

/// <summary>Projects applicable table-role defaults and explicitly selected modern text styles through the shared decoder.</summary>
internal sealed class IWorkTableTextStyleReader {
    private readonly IWorkSourceDocument _source;
    private readonly IWorkTableCellStyleCatalog _catalog;
    private readonly IWorkTextReader.TableStyleResolver _resolver;
    private readonly Dictionary<uint, IWorkParagraphStyle?> _selected = new();
    internal bool FullyReconstructed { get; private set; } = true;
    internal IWorkTableTextStyles Defaults { get; }

    internal IWorkTableTextStyleReader(IWorkSourceDocument source, IWorkArchiveRecord model,
        IWorkWireMessage message, IWorkTableCellStyleCatalog catalog, IWorkProjectionBudget budget,
        IWorkSourceReferenceIssueCollector references, int rows, int columns,
        int headerRows, int headerColumns, int footerRows) {
        _source = source; _catalog = catalog;
        _resolver = new IWorkTextReader.TableStyleResolver(source.Index, budget, references);
        // Row headers precede column headers; a column header also takes precedence in a footer row.
        bool hasCells = rows > 0 && columns > 0;
        Defaults = new IWorkTableTextStyles(
            ReadRole(24, hasCells && rows > headerRows + footerRows && columns > headerColumns),
            ReadRole(25, hasCells && headerRows > 0),
            ReadRole(26, hasCells && rows > headerRows && headerColumns > 0),
            ReadRole(27, hasCells && footerRows > 0 && columns > headerColumns));

        IWorkParagraphStyle? ReadRole(int field, bool applicable) {
            if (!applicable || !message.HasField(field)) return null;
            source.CancellationToken.ThrowIfCancellationRequested();
            IWorkArchiveRecord? record = references.ReadOne(model, message, field, allowedType: static type => type == 2022);
            bool complete = message.FieldCount(field) == 1
                && !message.HasUnexpectedWireKind(field, IWorkWireKind.Bytes)
                && record?.MessageType == 2022;
            IWorkParagraphStyle? style = complete ? _resolver.Read(record!.Identifier, ref complete) : null;
            if (!complete) {
                FullyReconstructed = false;
                references.Declarations.Record(model, field.ToString(System.Globalization.CultureInfo.InvariantCulture),
                    message.FieldCount(field), IWorkSourceDeclarationIssueKind.RejectedMessageSet);
            }
            return style;
        }
    }

    internal IWorkParagraphStyle? ReadSelected(uint key) {
        _source.CancellationToken.ThrowIfCancellationRequested();
        if (_selected.TryGetValue(key, out var cached)) return cached;
        bool complete = true;
        IWorkArchiveRecord? record = _catalog.ReadTextStyle(key, ref complete);
        IWorkParagraphStyle? style = record == null ? null : _resolver.Read(record.Identifier, ref complete);
        if (!complete) { FullyReconstructed = false; _catalog.RecordTextStyleFailure(key); }
        _selected.Add(key, style);
        return style;
    }
}
