namespace OfficeIMO.IWork.Internal;

/// <summary>Projects applicable table-role defaults and explicitly selected modern text styles through the shared decoder.</summary>
internal sealed class IWorkTableTextStyleReader {
    private readonly IWorkSourceDocument _source;
    private readonly IWorkProjectionBudget _budget;
    private readonly IWorkTableCellStyleCatalog _catalog;
    private readonly IWorkTextReader.TableStyleResolver _resolver;
    private readonly Dictionary<uint, IWorkParagraphStyle?> _selected = new();
    internal bool FullyReconstructed { get; private set; } = true;
    internal IWorkTableTextStyles Defaults { get; }

    internal IWorkTableTextStyleReader(IWorkSourceDocument source, IWorkArchiveRecord model,
        IWorkWireMessage message, IWorkTableCellStyleCatalog catalog, IWorkProjectionBudget budget,
        IWorkSourceReferenceIssueCollector references, int rows, int columns,
        int headerRows, int headerColumns, int footerRows) {
        _source = source; _catalog = catalog; _budget = budget;
        _resolver = new IWorkTextReader.TableStyleResolver(source.Index, budget, references);
        // Row headers precede column headers; a column header also takes precedence in a footer row.
        bool hasCells = rows > 0 && columns > 0;
        Defaults = new IWorkTableTextStyles(
            ReadRole(24, hasCells && rows > headerRows + footerRows && columns > headerColumns),
            ReadRole(25, hasCells && headerRows > 0),
            ReadRole(26, hasCells && rows > headerRows && headerColumns > 0),
            ReadRole(27, hasCells && footerRows > 0 && columns > headerColumns));

        ChargeTabs(Defaults.Body, (long)Math.Max(0, rows - headerRows - footerRows) * Math.Max(0, columns - headerColumns));
        ChargeTabs(Defaults.HeaderRow, (long)headerRows * columns);
        ChargeTabs(Defaults.HeaderColumn, (long)Math.Max(0, rows - headerRows) * headerColumns);
        ChargeTabs(Defaults.FooterRow, (long)footerRows * Math.Max(0, columns - headerColumns));

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
        if (_selected.TryGetValue(key, out var cached)) { ChargeTabs(cached, 1); return cached; }
        bool complete = true;
        IWorkArchiveRecord? record = _catalog.ReadTextStyle(key, ref complete);
        IWorkParagraphStyle? style = record == null ? null : _resolver.Read(record.Identifier, ref complete);
        if (!complete) { FullyReconstructed = false; _catalog.RecordTextStyleFailure(key); }
        _selected.Add(key, style);
        ChargeTabs(style, 1);
        return style;
    }
    internal void ChargeCellParagraphUses(IWorkTable table) {
        foreach (IWorkTableCell cell in table.Cells) {
            _source.CancellationToken.ThrowIfCancellationRequested();
            // Adapters apply defaults again to each populated paragraph, even
            // when a paragraph's explicit tabs subsequently replace them.
            ChargeTabs(table.GetParagraphStyle(cell.Row, cell.Column),
                Math.Max(1, cell.RichText?.Paragraphs.Count ?? 0));
        }
    }

    private void ChargeTabs(IWorkParagraphStyle? style, long uses) {
        int tabCount = style?.TabStops?.Count ?? 0;
        if (tabCount == 0) return;
        if (uses > int.MaxValue / tabCount) throw new InvalidDataException("Tab-stop use exceeds the text item projection limit.");
        _budget.AddTextItems(checked((int)uses * tabCount));
    }

}
