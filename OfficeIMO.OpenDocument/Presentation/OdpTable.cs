namespace OfficeIMO.OpenDocument;

/// <summary>An XML-backed presentation table frame.</summary>
public sealed class OdpTable : OdpShape {
    internal OdpTable(OdpPresentation presentation, XElement element) : base(presentation, element) { }
    private XElement TableElement => Element.Element(OdfNamespaces.Table + "table") ?? throw new InvalidDataException("ODP table frame has no table:table.");
    /// <summary>Table rows.</summary>
    public IReadOnlyList<OdpTableRow> Rows {
        get {
            List<XElement> elements = OdfTableRowElements.Enumerate(TableElement).ToList();
            return new OdfRepeatedElementCollection<OdpTableRow>(elements, OdfNamespaces.Table + "number-rows-repeated",
                (element, offset) => {
                    long logicalIndex = LogicalIndex(elements, element, offset, OdfNamespaces.Table + "number-rows-repeated");
                    return new OdpTableRow(Presentation, element, offset, () => ResolveRowElement(logicalIndex));
                });
        }
    }
    /// <summary>Gets a zero-based table cell.</summary>
    public OdpTableCell Cell(int row, int column) {
        if (row < 0) throw new ArgumentOutOfRangeException(nameof(row));
        if (column < 0) throw new ArgumentOutOfRangeException(nameof(column));
        OdpTableRow tableRow = Rows.ElementAtOrDefault(row) ?? throw new ArgumentOutOfRangeException(nameof(row));
        return tableRow.Cells.ElementAtOrDefault(column) ?? throw new ArgumentOutOfRangeException(nameof(column));
    }
    /// <summary>Merges a rectangular table range.</summary>
    public OdpTableCell Merge(int row, int column, int rowSpan, int columnSpan) {
        if (rowSpan < 1) throw new ArgumentOutOfRangeException(nameof(rowSpan));
        if (columnSpan < 1) throw new ArgumentOutOfRangeException(nameof(columnSpan));
        OdpTableCell anchor = Cell(row, column); anchor.SetSpans(rowSpan, columnSpan);
        for (int y = 0; y < rowSpan; y++) for (int x = 0; x < columnSpan; x++) if (x != 0 || y != 0) Cell(row + y, column + x).ReplaceWithCoveredCell();
        return anchor;
    }
    internal static OdpTable Create(OdpPresentation presentation, OdfRect bounds, int rows, int columns, string name) {
        if (rows < 1) throw new ArgumentOutOfRangeException(nameof(rows));
        if (columns < 1) throw new ArgumentOutOfRangeException(nameof(columns));
        var table = new XElement(OdfNamespaces.Table + "table", new XAttribute(OdfNamespaces.Table + "name", name));
        table.Add(new XElement(OdfNamespaces.Table + "table-column", new XAttribute(OdfNamespaces.Table + "number-columns-repeated", columns)));
        for (int row = 0; row < rows; row++) {
            var rowElement = new XElement(OdfNamespaces.Table + "table-row");
            for (int column = 0; column < columns; column++) rowElement.Add(OdpTableCell.CreateElement());
            table.Add(rowElement);
        }
        var frame = new XElement(OdfNamespaces.Draw + "frame", new XAttribute(OdfNamespaces.Draw + "name", name), table);
        ApplyBounds(frame, bounds); return new OdpTable(presentation, frame);
    }
    private OdfRepeatedElementPosition ResolveRowElement(long logicalIndex) => OdsRepeatModel.Resolve(
        OdfTableRowElements.Enumerate(TableElement).ToList(), OdfNamespaces.Table + "number-rows-repeated", logicalIndex);
    private static long LogicalIndex(IReadOnlyList<XElement> elements, XElement selected, long offset, XName repeatAttribute) {
        long index = 0;
        foreach (XElement element in elements) {
            if (ReferenceEquals(element, selected)) return checked(index + offset);
            index = checked(index + OdsRepeatModel.Read(element, repeatAttribute));
        }
        throw new InvalidOperationException("Repeated ODF element is no longer present in its collection.");
    }
}

/// <summary>An XML-backed presentation table row.</summary>
public sealed class OdpTableRow {
    private readonly OdpPresentation _presentation; private XElement _element; private long _repeatOffset;
    private Func<OdfRepeatedElementPosition>? _resolveRow;
    internal OdpTableRow(OdpPresentation presentation, XElement element, long repeatOffset = 0,
        Func<OdfRepeatedElementPosition>? resolveRow = null) {
        _presentation = presentation; _element = element; _repeatOffset = repeatOffset; _resolveRow = resolveRow;
    }
    /// <summary>Cells, including covered merged positions.</summary>
    public IReadOnlyList<OdpTableCell> Cells {
        get {
            List<XElement> elements = _element.Elements()
                .Where(element => element.Name == OdfNamespaces.Table + "table-cell" || element.Name == OdfNamespaces.Table + "covered-table-cell")
                .ToList();
            return new OdfRepeatedElementCollection<OdpTableCell>(elements, OdfNamespaces.Table + "number-columns-repeated",
                (element, offset) => {
                    long logicalIndex = LogicalIndex(elements, element, offset);
                    return new OdpTableCell(_presentation, element, offset, () => ResolveCellElement(logicalIndex));
                });
        }
    }
    private void EnsureMaterialized() {
        if (_resolveRow != null) {
            OdfRepeatedElementPosition position = _resolveRow();
            _element = position.Element;
            _repeatOffset = position.Offset;
            _resolveRow = null;
        }
        if (_element.Attribute(OdfNamespaces.Table + "number-rows-repeated") == null) return;
        _element = OdsRepeatModel.Split(_element, OdfNamespaces.Table + "number-rows-repeated", _repeatOffset);
    }
    private OdfRepeatedElementPosition ResolveCellElement(long logicalIndex) {
        EnsureMaterialized();
        return OdsRepeatModel.Resolve(_element.Elements()
            .Where(element => element.Name == OdfNamespaces.Table + "table-cell" || element.Name == OdfNamespaces.Table + "covered-table-cell")
            .ToList(), OdfNamespaces.Table + "number-columns-repeated", logicalIndex);
    }
    private static long LogicalIndex(IReadOnlyList<XElement> elements, XElement selected, long offset) {
        long index = 0;
        foreach (XElement element in elements) {
            if (ReferenceEquals(element, selected)) return checked(index + offset);
            index = checked(index + OdsRepeatModel.Read(element, OdfNamespaces.Table + "number-columns-repeated"));
        }
        throw new InvalidOperationException("Repeated ODF cell is no longer present in its row.");
    }
}

/// <summary>An XML-backed presentation table cell.</summary>
public sealed class OdpTableCell {
    private readonly OdpPresentation _presentation; private XElement _element; private long _repeatOffset;
    private Func<OdfRepeatedElementPosition>? _resolveRowCell;
    internal OdpTableCell(OdpPresentation presentation, XElement element, long repeatOffset = 0,
        Func<OdfRepeatedElementPosition>? resolveRowCell = null) {
        _presentation = presentation; _element = element; _repeatOffset = repeatOffset; _resolveRowCell = resolveRowCell;
    }
    /// <summary>True when covered by a merged cell.</summary>
    public bool IsCovered => _element.Name == OdfNamespaces.Table + "covered-table-cell";
    /// <summary>Number of rows spanned by a merged-cell anchor.</summary>
    public int RowSpan => ReadSpan(OdfNamespaces.Table + "number-rows-spanned");
    /// <summary>Number of columns spanned by a merged-cell anchor.</summary>
    public int ColumnSpan => ReadSpan(OdfNamespaces.Table + "number-columns-spanned");
    /// <summary>Decoded cell text.</summary>
    public string Text {
        get => string.Join("\n", _element.Elements(OdfNamespaces.Text + "p").Select(OdfTextCodec.Read));
        set {
            if (IsCovered) throw new InvalidOperationException("Covered table cells cannot contain text.");
            EnsureMaterialized();
            _element.RemoveNodes(); var paragraph = new XElement(OdfNamespaces.Text + "p"); OdfTextCodec.Append(paragraph, value); _element.Add(paragraph); Dirty();
        }
    }
    internal static XElement CreateElement() => new XElement(OdfNamespaces.Table + "table-cell", new XElement(OdfNamespaces.Text + "p"));
    internal void SetSpans(int rows, int columns) {
        EnsureMaterialized();
        _element.SetAttributeValue(OdfNamespaces.Table + "number-rows-spanned", rows > 1 ? rows : (int?)null);
        _element.SetAttributeValue(OdfNamespaces.Table + "number-columns-spanned", columns > 1 ? columns : (int?)null); Dirty();
    }
    internal void ReplaceWithCoveredCell() { EnsureMaterialized(); var covered = new XElement(OdfNamespaces.Table + "covered-table-cell"); _element.ReplaceWith(covered); _element = covered; Dirty(); }
    private int ReadSpan(XName name) {
        string? lexical = (string?)_element.Attribute(name);
        return int.TryParse(lexical, NumberStyles.None, CultureInfo.InvariantCulture, out int value) && value > 0 ? value : 1;
    }
    private void EnsureMaterialized() {
        if (_resolveRowCell != null) {
            OdfRepeatedElementPosition position = _resolveRowCell();
            _element = position.Element;
            _repeatOffset = position.Offset;
            _resolveRowCell = null;
        }
        if (_element.Attribute(OdfNamespaces.Table + "number-columns-repeated") == null) return;
        _element = OdsRepeatModel.Split(_element, OdfNamespaces.Table + "number-columns-repeated", _repeatOffset);
    }
    private void Dirty() => _presentation.MarkPartDirty("content.xml");
}
