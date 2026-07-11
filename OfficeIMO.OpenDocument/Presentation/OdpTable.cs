namespace OfficeIMO.OpenDocument;

/// <summary>An XML-backed presentation table frame.</summary>
public sealed class OdpTable : OdpShape {
    internal OdpTable(OdpPresentation presentation, XElement element) : base(presentation, element) { }
    private XElement TableElement => Element.Element(OdfNamespaces.Table + "table") ?? throw new InvalidDataException("ODP table frame has no table:table.");
    /// <summary>Table rows.</summary>
    public IReadOnlyList<OdpTableRow> Rows => new OdfRepeatedElementCollection<OdpTableRow>(
        OdfTableRowElements.Enumerate(TableElement).ToList(), OdfNamespaces.Table + "number-rows-repeated",
        (element, _) => new OdpTableRow(Presentation, element));
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
}

/// <summary>An XML-backed presentation table row.</summary>
public sealed class OdpTableRow {
    private readonly OdpPresentation _presentation; private readonly XElement _element;
    internal OdpTableRow(OdpPresentation presentation, XElement element) { _presentation = presentation; _element = element; }
    /// <summary>Cells, including covered merged positions.</summary>
    public IReadOnlyList<OdpTableCell> Cells => new OdfRepeatedElementCollection<OdpTableCell>(_element.Elements()
        .Where(element => element.Name == OdfNamespaces.Table + "table-cell" || element.Name == OdfNamespaces.Table + "covered-table-cell")
        .ToList(), OdfNamespaces.Table + "number-columns-repeated",
        (element, offset) => new OdpTableCell(_presentation, element, offset));
}

/// <summary>An XML-backed presentation table cell.</summary>
public sealed class OdpTableCell {
    private readonly OdpPresentation _presentation; private XElement _element; private readonly long _repeatOffset;
    internal OdpTableCell(OdpPresentation presentation, XElement element, long repeatOffset = 0) {
        _presentation = presentation; _element = element; _repeatOffset = repeatOffset;
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
        if (_element.Attribute(OdfNamespaces.Table + "number-columns-repeated") == null) return;
        _element = OdsRepeatModel.Split(_element, OdfNamespaces.Table + "number-columns-repeated", _repeatOffset);
    }
    private void Dirty() => _presentation.MarkPartDirty("content.xml");
}
