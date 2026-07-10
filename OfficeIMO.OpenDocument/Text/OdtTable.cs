namespace OfficeIMO.OpenDocument;

/// <summary>An XML-backed ODT table.</summary>
public sealed class OdtTable {
    private readonly OdtDocument _document;
    private readonly XElement _element;

    internal OdtTable(OdtDocument document, XElement element) {
        _document = document;
        _element = element;
    }

    /// <summary>Table name.</summary>
    public string Name {
        get => (string?)_element.Attribute(OdfNamespaces.Table + "name") ?? string.Empty;
        set {
            if (string.IsNullOrWhiteSpace(value)) throw new ArgumentException("Table name cannot be empty.", nameof(value));
            _element.SetAttributeValue(OdfNamespaces.Table + "name", value);
            Dirty();
        }
    }

    /// <summary>Rows in source order.</summary>
    public IReadOnlyList<OdtTableRow> Rows => _element.Elements(OdfNamespaces.Table + "table-row")
        .Select(element => new OdtTableRow(_document, element)).ToList();

    /// <summary>Gets a zero-based cell.</summary>
    public OdtTableCell Cell(int row, int column) {
        if (row < 0) throw new ArgumentOutOfRangeException(nameof(row));
        if (column < 0) throw new ArgumentOutOfRangeException(nameof(column));
        OdtTableRow tableRow = Rows.ElementAtOrDefault(row) ?? throw new ArgumentOutOfRangeException(nameof(row));
        return tableRow.Cells.ElementAtOrDefault(column) ?? throw new ArgumentOutOfRangeException(nameof(column));
    }

    /// <summary>Adds a row with the inferred or supplied column count.</summary>
    public OdtTableRow AddRow(int? columns = null) {
        int count = columns ?? (Rows.FirstOrDefault()?.Cells.Count ?? 1);
        if (count < 1) throw new ArgumentOutOfRangeException(nameof(columns));
        var row = new XElement(OdfNamespaces.Table + "table-row");
        for (int index = 0; index < count; index++) row.Add(OdtTableCell.CreateElement());
        _element.Add(row);
        Dirty();
        return new OdtTableRow(_document, row);
    }

    /// <summary>Merges a rectangular range and emits covered cells for non-anchor positions.</summary>
    public OdtTableCell Merge(int row, int column, int rowSpan, int columnSpan) {
        if (rowSpan < 1) throw new ArgumentOutOfRangeException(nameof(rowSpan));
        if (columnSpan < 1) throw new ArgumentOutOfRangeException(nameof(columnSpan));
        OdtTableCell anchor = Cell(row, column);
        anchor.SetSpans(rowSpan, columnSpan);
        for (int rowOffset = 0; rowOffset < rowSpan; rowOffset++) {
            for (int columnOffset = 0; columnOffset < columnSpan; columnOffset++) {
                if (rowOffset == 0 && columnOffset == 0) continue;
                OdtTableCell cell = Cell(row + rowOffset, column + columnOffset);
                cell.ReplaceWithCoveredCell();
            }
        }
        Dirty();
        return anchor;
    }

    internal XElement Element => _element;
    private void Dirty() => _document.MarkPartDirty("content.xml");
}

/// <summary>An XML-backed ODT table row.</summary>
public sealed class OdtTableRow {
    private readonly OdtDocument _document;
    private readonly XElement _element;

    internal OdtTableRow(OdtDocument document, XElement element) {
        _document = document;
        _element = element;
    }

    /// <summary>Cells, including covered cells, in source order.</summary>
    public IReadOnlyList<OdtTableCell> Cells => new OdfRepeatedElementCollection<OdtTableCell>(_element.Elements()
        .Where(element => element.Name == OdfNamespaces.Table + "table-cell" || element.Name == OdfNamespaces.Table + "covered-table-cell")
        .ToList(), OdfNamespaces.Table + "number-columns-repeated",
        (element, offset) => new OdtTableCell(_document, element, offset));

    /// <summary>Adds a cell.</summary>
    public OdtTableCell AddCell(string? text = null) {
        XElement cell = OdtTableCell.CreateElement(text);
        _element.Add(cell);
        _document.MarkPartDirty("content.xml");
        return new OdtTableCell(_document, cell);
    }
}

/// <summary>An XML-backed ODT table cell.</summary>
public sealed class OdtTableCell {
    private readonly OdtDocument _document;
    private XElement _element;
    private readonly long _repeatOffset;

    internal OdtTableCell(OdtDocument document, XElement element, long repeatOffset = 0) {
        _document = document;
        _element = element;
        _repeatOffset = repeatOffset;
    }

    /// <summary>True when this is a covered position in a merged range.</summary>
    public bool IsCovered => _element.Name == OdfNamespaces.Table + "covered-table-cell";
    /// <summary>Row span on the anchor cell.</summary>
    public int RowSpan => ReadCount(OdfNamespaces.Table + "number-rows-spanned");
    /// <summary>Column span on the anchor cell.</summary>
    public int ColumnSpan => ReadCount(OdfNamespaces.Table + "number-columns-spanned");
    /// <summary>Paragraphs directly stored in this cell.</summary>
    public IReadOnlyList<OdtParagraph> Paragraphs => _element.Elements()
        .Where(element => element.Name == OdfNamespaces.Text + "p" || element.Name == OdfNamespaces.Text + "h")
        .Select(element => new OdtParagraph(_document, element)).ToList();
    /// <summary>Cell text joined across paragraphs.</summary>
    public string Text {
        get => string.Join("\n", Paragraphs.Select(paragraph => paragraph.Text));
        set {
            if (IsCovered) throw new InvalidOperationException("Covered table cells cannot contain text.");
            EnsureMaterialized();
            _element.RemoveNodes();
            var paragraph = new XElement(OdfNamespaces.Text + "p");
            OdfTextCodec.Append(paragraph, value);
            _element.Add(paragraph);
            _element.SetAttributeValue(OdfNamespaces.Office + "value-type", "string");
            Dirty();
        }
    }

    /// <summary>Adds a paragraph to the cell.</summary>
    public OdtParagraph AddParagraph(string? text = null) {
        if (IsCovered) throw new InvalidOperationException("Covered table cells cannot contain paragraphs.");
        EnsureMaterialized();
        var paragraph = new XElement(OdfNamespaces.Text + "p");
        OdfTextCodec.Append(paragraph, text);
        _element.Add(paragraph);
        Dirty();
        return new OdtParagraph(_document, paragraph);
    }

    internal static XElement CreateElement(string? text = null) {
        var paragraph = new XElement(OdfNamespaces.Text + "p");
        OdfTextCodec.Append(paragraph, text);
        return new XElement(OdfNamespaces.Table + "table-cell",
            new XAttribute(OdfNamespaces.Office + "value-type", "string"), paragraph);
    }

    internal void SetSpans(int rows, int columns) {
        EnsureMaterialized();
        _element.SetAttributeValue(OdfNamespaces.Table + "number-rows-spanned", rows > 1 ? rows : (int?)null);
        _element.SetAttributeValue(OdfNamespaces.Table + "number-columns-spanned", columns > 1 ? columns : (int?)null);
        Dirty();
    }

    internal void ReplaceWithCoveredCell() {
        EnsureMaterialized();
        var covered = new XElement(OdfNamespaces.Table + "covered-table-cell");
        _element.ReplaceWith(covered);
        _element = covered;
        Dirty();
    }

    private int ReadCount(XName name) {
        return int.TryParse((string?)_element.Attribute(name), NumberStyles.Integer, CultureInfo.InvariantCulture, out int value) && value > 0 ? value : 1;
    }

    private void EnsureMaterialized() {
        if (_element.Attribute(OdfNamespaces.Table + "number-columns-repeated") == null) return;
        _element = OdsRepeatModel.Split(_element, OdfNamespaces.Table + "number-columns-repeated", _repeatOffset);
    }

    private void Dirty() => _document.MarkPartDirty("content.xml");
}
