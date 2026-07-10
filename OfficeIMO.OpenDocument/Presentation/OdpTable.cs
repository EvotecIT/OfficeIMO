namespace OfficeIMO.OpenDocument;

/// <summary>An XML-backed presentation table frame.</summary>
public sealed class OdpTable : OdpShape {
    internal OdpTable(OdpPresentation presentation, XElement element) : base(presentation, element) { }
    private XElement TableElement => Element.Element(OdfNamespaces.Table + "table") ?? throw new InvalidDataException("ODP table frame has no table:table.");
    /// <summary>Table rows.</summary>
    public IReadOnlyList<OdpTableRow> Rows => TableElement.Elements(OdfNamespaces.Table + "table-row")
        .Select(element => new OdpTableRow(Presentation, element)).ToList();
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
    public IReadOnlyList<OdpTableCell> Cells => _element.Elements()
        .Where(element => element.Name == OdfNamespaces.Table + "table-cell" || element.Name == OdfNamespaces.Table + "covered-table-cell")
        .Select(element => new OdpTableCell(_presentation, element)).ToList();
}

/// <summary>An XML-backed presentation table cell.</summary>
public sealed class OdpTableCell {
    private readonly OdpPresentation _presentation; private XElement _element;
    internal OdpTableCell(OdpPresentation presentation, XElement element) { _presentation = presentation; _element = element; }
    /// <summary>True when covered by a merged cell.</summary>
    public bool IsCovered => _element.Name == OdfNamespaces.Table + "covered-table-cell";
    /// <summary>Decoded cell text.</summary>
    public string Text {
        get => string.Join("\n", _element.Elements(OdfNamespaces.Text + "p").Select(OdfTextCodec.Read));
        set {
            if (IsCovered) throw new InvalidOperationException("Covered table cells cannot contain text.");
            _element.RemoveNodes(); var paragraph = new XElement(OdfNamespaces.Text + "p"); OdfTextCodec.Append(paragraph, value); _element.Add(paragraph); Dirty();
        }
    }
    internal static XElement CreateElement() => new XElement(OdfNamespaces.Table + "table-cell", new XElement(OdfNamespaces.Text + "p"));
    internal void SetSpans(int rows, int columns) {
        _element.SetAttributeValue(OdfNamespaces.Table + "number-rows-spanned", rows > 1 ? rows : (int?)null);
        _element.SetAttributeValue(OdfNamespaces.Table + "number-columns-spanned", columns > 1 ? columns : (int?)null); Dirty();
    }
    internal void ReplaceWithCoveredCell() { var covered = new XElement(OdfNamespaces.Table + "covered-table-cell"); _element.ReplaceWith(covered); _element = covered; Dirty(); }
    private void Dirty() => _presentation.MarkPartDirty("content.xml");
}
