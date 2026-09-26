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
    public IReadOnlyList<OdtTableRow> Rows {
        get {
            List<XElement> elements = OdfTableRowElements.Enumerate(_element).ToList();
            return new OdfRepeatedElementCollection<OdtTableRow>(elements, OdfNamespaces.Table + "number-rows-repeated",
                (element, offset) => {
                    long logicalIndex = LogicalIndex(elements, element, offset, OdfNamespaces.Table + "number-rows-repeated");
                    return new OdtTableRow(_document, element, offset, () => ResolveRowElement(logicalIndex));
                });
        }
    }

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
        // Resolve and validate every logical cell before changing the anchor.
        // Materializing a repeated row or cell that already contains a note
        // would clone its identity, so the merge must fail without mutation.
        for (int rowOffset = 0; rowOffset < rowSpan; rowOffset++) {
            for (int columnOffset = 0; columnOffset < columnSpan; columnOffset++) {
                Cell(row + rowOffset, column + columnOffset).PreflightMerge();
            }
        }
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
    private OdfRepeatedElementPosition ResolveRowElement(long logicalIndex) => OdsRepeatModel.Resolve(
        OdfTableRowElements.Enumerate(_element).ToList(), OdfNamespaces.Table + "number-rows-repeated", logicalIndex);
    private static long LogicalIndex(IReadOnlyList<XElement> elements, XElement selected, long offset, XName repeatAttribute) {
        long index = 0;
        foreach (XElement element in elements) {
            if (ReferenceEquals(element, selected)) return checked(index + offset);
            index = checked(index + OdsRepeatModel.Read(element, repeatAttribute));
        }
        throw new InvalidOperationException("Repeated ODF element is no longer present in its collection.");
    }
    private void Dirty() => _document.MarkPartDirty("content.xml");
}

/// <summary>An XML-backed ODT table row.</summary>
public sealed class OdtTableRow {
    private readonly OdtDocument _document;
    private XElement _element;
    private long _repeatOffset;
    private Func<OdfRepeatedElementPosition>? _resolveRow;

    internal OdtTableRow(OdtDocument document, XElement element, long repeatOffset = 0,
        Func<OdfRepeatedElementPosition>? resolveRow = null) {
        _document = document;
        _element = element;
        _repeatOffset = repeatOffset;
        _resolveRow = resolveRow;
    }

    /// <summary>Cells, including covered cells, in source order.</summary>
    public IReadOnlyList<OdtTableCell> Cells {
        get {
            List<XElement> elements = _element.Elements()
                .Where(element => element.Name == OdfNamespaces.Table + "table-cell" || element.Name == OdfNamespaces.Table + "covered-table-cell")
                .ToList();
            return new OdfRepeatedElementCollection<OdtTableCell>(elements, OdfNamespaces.Table + "number-columns-repeated",
                (element, offset) => {
                    long logicalIndex = LogicalIndex(elements, element, offset);
                    return new OdtTableCell(_document, element, offset, () => ResolveCellElement(logicalIndex));
                });
        }
    }

    /// <summary>Adds a cell.</summary>
    public OdtTableCell AddCell(string? text = null) {
        EnsureMaterialized();
        XElement cell = OdtTableCell.CreateElement(text);
        _element.Add(cell);
        _document.MarkPartDirty("content.xml");
        return new OdtTableCell(_document, cell);
    }

    private void EnsureMaterialized() {
        if (_resolveRow != null) {
            OdfRepeatedElementPosition position = _resolveRow();
            _element = position.Element;
            _repeatOffset = position.Offset;
            _resolveRow = null;
        }
        if (_element.Attribute(OdfNamespaces.Table + "number-rows-repeated") == null) return;
        if (OdsRepeatModel.Read(_element, OdfNamespaces.Table + "number-rows-repeated") > 1 &&
            _element.Descendants(OdfNamespaces.Text + "note").Any()) {
            throw new NotSupportedException("Splitting a repeated table row containing a note is not supported.");
        }
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

/// <summary>An XML-backed ODT table cell.</summary>
public sealed class OdtTableCell {
    private readonly OdtDocument _document;
    private XElement _element;
    private long _repeatOffset;
    private Func<OdfRepeatedElementPosition>? _resolveRowCell;

    internal OdtTableCell(OdtDocument document, XElement element, long repeatOffset = 0,
        Func<OdfRepeatedElementPosition>? resolveRowCell = null) {
        _document = document;
        _element = element;
        _repeatOffset = repeatOffset;
        _resolveRowCell = resolveRowCell;
    }

    /// <summary>True when this is a covered position in a merged range.</summary>
    public bool IsCovered => _element.Name == OdfNamespaces.Table + "covered-table-cell";
    /// <summary>Row span on the anchor cell.</summary>
    public int RowSpan => ReadCount(OdfNamespaces.Table + "number-rows-spanned");
    /// <summary>Column span on the anchor cell.</summary>
    public int ColumnSpan => ReadCount(OdfNamespaces.Table + "number-columns-spanned");
    /// <summary>Paragraphs directly stored in this cell.</summary>
    public IReadOnlyList<OdtParagraph> Paragraphs {
        get {
            // Reading a repeated cell must remain sparse. A paragraph resolves its
            // logical row and cell only if a note is actually inserted.
            return _element.Elements()
                .Where(element => element.Name == OdfNamespaces.Text + "p" || element.Name == OdfNamespaces.Text + "h")
                .Select((element, index) => new OdtParagraph(_document, element,
                    materializeForNote: () => ResolveParagraphForNote(index))).ToList();
        }
    }

    private XElement ResolveParagraphForNote(int index) {
        // Splitting a repeated row or cell clones its existing notes. Their IDs
        // would then be shared by several physical notes, so reject before the
        // first XML mutation instead of inserting against a stale note index.
        bool repeatsCellWithNote =
            OdsRepeatModel.Read(_element, OdfNamespaces.Table + "number-columns-repeated") > 1 &&
            _element.Descendants(OdfNamespaces.Text + "note").Any();
        bool repeatsRowWithNote = _element.Parent is XElement row &&
            row.Name == OdfNamespaces.Table + "table-row" &&
            OdsRepeatModel.Read(row, OdfNamespaces.Table + "number-rows-repeated") > 1 &&
            row.Descendants(OdfNamespaces.Text + "note").Any();
        if (repeatsCellWithNote || repeatsRowWithNote) {
            throw new NotSupportedException("Adding a note while splitting a repeated table cell or row containing a note is not supported.");
        }
        EnsureMaterialized();
        return _element.Elements()
            .Where(element => element.Name == OdfNamespaces.Text + "p" || element.Name == OdfNamespaces.Text + "h")
            .ElementAt(index);
    }
    /// <summary>Cell text joined across paragraphs.</summary>
    public string Text {
        get => string.Join("\n", _element.Elements()
            .Where(element => element.Name == OdfNamespaces.Text + "p" || element.Name == OdfNamespaces.Text + "h")
            .Select(element => new OdtParagraph(_document, element).Text));
        set {
            if (IsCovered) throw new InvalidOperationException("Covered table cells cannot contain text.");
            EnsureMaterialized();
            bool hadNotes = _element.Descendants(OdfNamespaces.Text + "note").Any();
            if (hadNotes) _document.PrepareNoteIndexForMutation();
            _element.RemoveNodes();
            var paragraph = new XElement(OdfNamespaces.Text + "p");
            OdfTextCodec.Append(paragraph, value);
            _element.Add(paragraph);
            _element.SetAttributeValue(OdfNamespaces.Office + "value-type", "string");
            if (hadNotes) _document.RefreshNoteIndexAfterMutation();
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
        bool hadNotes = _element.Descendants(OdfNamespaces.Text + "note").Any();
        if (hadNotes) _document.PrepareNoteIndexForMutation();
        var covered = new XElement(OdfNamespaces.Table + "covered-table-cell");
        _element.ReplaceWith(covered);
        _element = covered;
        if (hadNotes) _document.RefreshNoteIndexAfterMutation();
        Dirty();
    }

    private int ReadCount(XName name) {
        return int.TryParse((string?)_element.Attribute(name), NumberStyles.Integer, CultureInfo.InvariantCulture, out int value) && value > 0 ? value : 1;
    }

    private void EnsureMaterialized() {
        bool resolvedRow = false;
        if (_resolveRowCell != null) {
            OdfRepeatedElementPosition position = _resolveRowCell();
            _element = position.Element;
            _repeatOffset = position.Offset;
            _resolveRowCell = null;
            resolvedRow = true;
        }
        if (_element.Attribute(OdfNamespaces.Table + "number-columns-repeated") != null) {
            if (OdsRepeatModel.Read(_element, OdfNamespaces.Table + "number-columns-repeated") > 1 &&
                _element.Descendants(OdfNamespaces.Text + "note").Any()) {
                throw new NotSupportedException("Splitting a repeated table cell containing a note is not supported.");
            }
            _element = OdsRepeatModel.Split(_element, OdfNamespaces.Table + "number-columns-repeated", _repeatOffset);
            Dirty();
        } else if (resolvedRow) {
            Dirty();
        }
    }

    internal void PreflightMerge() {
        if (OdsRepeatModel.Read(_element, OdfNamespaces.Table + "number-columns-repeated") > 1 &&
            _element.Descendants(OdfNamespaces.Text + "note").Any() ||
            _element.Parent is XElement row && row.Name == OdfNamespaces.Table + "table-row" &&
            OdsRepeatModel.Read(row, OdfNamespaces.Table + "number-rows-repeated") > 1 &&
            row.Descendants(OdfNamespaces.Text + "note").Any()) {
            throw new NotSupportedException("Merging a repeated table row or cell containing a note is not supported.");
        }
    }

    private void Dirty() => _document.MarkPartDirty("content.xml");
}
