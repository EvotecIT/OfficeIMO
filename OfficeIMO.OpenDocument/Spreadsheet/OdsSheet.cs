namespace OfficeIMO.OpenDocument;

/// <summary>An XML-backed ODS worksheet with sparse repeat-run editing.</summary>
public sealed class OdsSheet {
    /// <summary>Default maximum number of cells that one merge operation may materialize.</summary>
    public const long DefaultMaximumMergeCells = 100_000;

    private readonly OdsDocument _document;

    internal OdsSheet(OdsDocument document, XElement element) { _document = document; Element = element; }

    /// <summary>Worksheet name.</summary>
    public string Name {
        get => (string?)Element.Attribute(OdfNamespaces.Table + "name") ?? string.Empty;
        set {
            if (string.IsNullOrWhiteSpace(value)) throw new ArgumentException("Worksheet name cannot be empty.", nameof(value));
            if (_document.Sheets.Any(sheet => !ReferenceEquals(sheet.Element, Element) && string.Equals(sheet.Name, value, StringComparison.Ordinal))) {
                throw new InvalidOperationException($"A worksheet named '{value}' already exists.");
            }
            Element.SetAttributeValue(OdfNamespaces.Table + "name", value); Dirty();
        }
    }

    /// <summary>Whether the sheet is hidden.</summary>
    public bool Hidden {
        get => (string?)Element.Attribute(OdfNamespaces.Table + "visibility") == "collapse";
        set { Element.SetAttributeValue(OdfNamespaces.Table + "visibility", value ? "collapse" : null); Dirty(); }
    }

    /// <summary>Optional ODF print range expression.</summary>
    public string? PrintRanges {
        get => (string?)Element.Attribute(OdfNamespaces.Table + "print-ranges");
        set { Element.SetAttributeValue(OdfNamespaces.Table + "print-ranges", value); Dirty(); }
    }

    /// <summary>Sparse row runs without expanding <c>table:number-rows-repeated</c>.</summary>
    public IReadOnlyList<OdsRowRun> RowRuns {
        get {
            var runs = new List<OdsRowRun>();
            long start = 0;
            foreach (XElement row in RowElements()) {
                long repeat = OdsRepeatModel.Read(row, OdfNamespaces.Table + "number-rows-repeated");
                runs.Add(new OdsRowRun(_document, row, start, repeat));
                start = checked(start + repeat);
            }
            return runs;
        }
    }

    /// <summary>Sparse column definition runs without expanding repeats.</summary>
    public IReadOnlyList<OdsColumnRun> ColumnRuns {
        get {
            var runs = new List<OdsColumnRun>();
            long start = 0;
            foreach (XElement column in Element.Elements(OdfNamespaces.Table + "table-column")) {
                long repeat = OdsRepeatModel.Read(column, OdfNamespaces.Table + "number-columns-repeated");
                runs.Add(new OdsColumnRun(_document, column, start, repeat));
                start = checked(start + repeat);
            }
            return runs;
        }
    }

    /// <summary>Logical row count represented by the sparse run model.</summary>
    public long RowCount => RowRuns.Count == 0 ? 0 : checked(RowRuns[RowRuns.Count - 1].StartRow + RowRuns[RowRuns.Count - 1].RepeatCount);

    /// <summary>Smallest rectangle containing cells with a value, formula, or text.</summary>
    public OdsUsedRange? UsedRange {
        get {
            long rowStart = 0;
            long? firstRow = null, firstColumn = null, lastRow = null, lastColumn = null;
            foreach (XElement row in RowElements()) {
                long rowRepeat = OdsRepeatModel.Read(row, OdfNamespaces.Table + "number-rows-repeated");
                long columnStart = 0;
                foreach (XElement cell in CellElements(row)) {
                    long cellRepeat = OdsRepeatModel.Read(cell, OdfNamespaces.Table + "number-columns-repeated");
                    if (!OdsCell.IsEmpty(cell)) {
                        firstRow = !firstRow.HasValue ? rowStart : Math.Min(firstRow.Value, rowStart);
                        firstColumn = !firstColumn.HasValue ? columnStart : Math.Min(firstColumn.Value, columnStart);
                        lastRow = Math.Max(lastRow ?? rowStart, checked(rowStart + rowRepeat - 1));
                        lastColumn = Math.Max(lastColumn ?? columnStart, checked(columnStart + cellRepeat - 1));
                    }
                    columnStart = checked(columnStart + cellRepeat);
                }
                rowStart = checked(rowStart + rowRepeat);
            }
            return firstRow.HasValue ? new OdsUsedRange(firstRow.Value, firstColumn!.Value, lastRow!.Value, lastColumn!.Value) : (OdsUsedRange?)null;
        }
    }

    /// <summary>Gets an editable zero-based cell, splitting only the containing row and cell runs.</summary>
    public OdsCell Cell(long row, long column) {
        if (row < 0) throw new ArgumentOutOfRangeException(nameof(row));
        if (column < 0) throw new ArgumentOutOfRangeException(nameof(column));
        XElement rowElement = GetRowForEdit(row);
        XElement cellElement = GetCellForEdit(rowElement, column);
        return new OdsCell(_document, cellElement);
    }

    /// <summary>Gets an editable zero-based row, splitting its repeat run without expanding it.</summary>
    public OdsRow Row(long row) {
        if (row < 0) throw new ArgumentOutOfRangeException(nameof(row));
        return new OdsRow(_document, GetRowForEdit(row));
    }

    /// <summary>Gets an editable zero-based column definition, creating a sparse definition when needed.</summary>
    public OdsColumn Column(long column) {
        if (column < 0) throw new ArgumentOutOfRangeException(nameof(column));
        long start = 0;
        foreach (XElement element in Element.Elements(OdfNamespaces.Table + "table-column").ToList()) {
            long count = OdsRepeatModel.Read(element, OdfNamespaces.Table + "number-columns-repeated");
            if (column < checked(start + count)) {
                XElement target = OdsRepeatModel.Split(element, OdfNamespaces.Table + "number-columns-repeated", column - start);
                Dirty();
                return new OdsColumn(_document, target);
            }
            start = checked(start + count);
        }
        long required = checked(column - start + 1);
        var added = new XElement(OdfNamespaces.Table + "table-column");
        OdsRepeatModel.Set(added, OdfNamespaces.Table + "number-columns-repeated", required);
        XElement? firstRow = RowElements().FirstOrDefault();
        XElement? insertionPoint = firstRow?.Parent?.Name == OdfNamespaces.Table + "table-header-rows"
            ? firstRow.Parent
            : firstRow;
        if (insertionPoint == null) Element.Add(added); else insertionPoint.AddBeforeSelf(added);
        XElement result = OdsRepeatModel.Split(added, OdfNamespaces.Table + "number-columns-repeated", required - 1);
        Dirty();
        return new OdsColumn(_document, result);
    }

    /// <summary>Reads a value without splitting or expanding repeat runs.</summary>
    public OdsCellValue GetValue(long row, long column) {
        if (row < 0) throw new ArgumentOutOfRangeException(nameof(row));
        if (column < 0) throw new ArgumentOutOfRangeException(nameof(column));
        XElement? rowElement = FindPrototypeRow(row);
        if (rowElement == null) return OdsCellValue.Empty;
        long start = 0;
        foreach (XElement cell in CellElements(rowElement)) {
            long count = OdsRepeatModel.Read(cell, OdfNamespaces.Table + "number-columns-repeated");
            if (column < checked(start + count)) return OdsCell.ReadValue(cell);
            start = checked(start + count);
        }
        return OdsCellValue.Empty;
    }

    /// <summary>Reads a formula without splitting or expanding repeat runs.</summary>
    public string? GetFormula(long row, long column) {
        if (row < 0) throw new ArgumentOutOfRangeException(nameof(row));
        if (column < 0) throw new ArgumentOutOfRangeException(nameof(column));
        XElement? rowElement = FindPrototypeRow(row);
        if (rowElement == null) return null;
        long start = 0;
        foreach (XElement cell in CellElements(rowElement)) {
            long count = OdsRepeatModel.Read(cell, OdfNamespaces.Table + "number-columns-repeated");
            if (column < checked(start + count)) return (string?)cell.Attribute(OdfNamespaces.Table + "formula");
            start = checked(start + count);
        }
        return null;
    }

    /// <summary>Merges a rectangular cell range and marks non-anchor positions as covered cells.</summary>
    public OdsCell Merge(long row, long column, long rowSpan, long columnSpan) {
        return Merge(row, column, rowSpan, columnSpan, DefaultMaximumMergeCells);
    }

    /// <summary>Merges a rectangular cell range under an explicit materialization bound.</summary>
    public OdsCell Merge(long row, long column, long rowSpan, long columnSpan, long maximumMaterializedCells) {
        if (row < 0) throw new ArgumentOutOfRangeException(nameof(row));
        if (column < 0) throw new ArgumentOutOfRangeException(nameof(column));
        if (rowSpan < 1) throw new ArgumentOutOfRangeException(nameof(rowSpan));
        if (columnSpan < 1) throw new ArgumentOutOfRangeException(nameof(columnSpan));
        if (maximumMaterializedCells < 1) throw new ArgumentOutOfRangeException(nameof(maximumMaterializedCells));
        long mergeCells;
        try {
            mergeCells = checked(rowSpan * columnSpan);
            _ = checked(row + rowSpan - 1);
            _ = checked(column + columnSpan - 1);
        } catch (OverflowException) {
            throw new ArgumentOutOfRangeException(nameof(rowSpan), "Merge dimensions exceed the supported coordinate range.");
        }
        if (mergeCells > maximumMaterializedCells) {
            throw new InvalidOperationException($"Merge would materialize {mergeCells} cells, exceeding the configured limit of {maximumMaterializedCells}.");
        }
        OdsCell anchor = Cell(row, column);
        anchor.SetSpans(rowSpan, columnSpan);
        for (long rowOffset = 0; rowOffset < rowSpan; rowOffset++) {
            for (long columnOffset = 0; columnOffset < columnSpan; columnOffset++) {
                if (rowOffset == 0 && columnOffset == 0) continue;
                Cell(checked(row + rowOffset), checked(column + columnOffset)).ReplaceWithCoveredCell();
            }
        }
        return anchor;
    }

    internal XElement Element { get; }

    private XElement GetRowForEdit(long rowIndex) {
        long start = 0;
        foreach (XElement element in RowElements().ToList()) {
            long count = OdsRepeatModel.Read(element, OdfNamespaces.Table + "number-rows-repeated");
            if (rowIndex < checked(start + count)) {
                XElement target = OdsRepeatModel.Split(element, OdfNamespaces.Table + "number-rows-repeated", rowIndex - start);
                Dirty();
                return target;
            }
            start = checked(start + count);
        }
        long required = checked(rowIndex - start + 1);
        var added = new XElement(OdfNamespaces.Table + "table-row", new XElement(OdfNamespaces.Table + "table-cell"));
        OdsRepeatModel.Set(added, OdfNamespaces.Table + "number-rows-repeated", required);
        Element.Add(added);
        XElement result = OdsRepeatModel.Split(added, OdfNamespaces.Table + "number-rows-repeated", required - 1);
        Dirty();
        return result;
    }

    private XElement? FindPrototypeRow(long rowIndex) {
        long start = 0;
        foreach (XElement element in RowElements()) {
            long count = OdsRepeatModel.Read(element, OdfNamespaces.Table + "number-rows-repeated");
            if (rowIndex < checked(start + count)) return element;
            start = checked(start + count);
        }
        return null;
    }

    private XElement GetCellForEdit(XElement row, long columnIndex) {
        long start = 0;
        foreach (XElement element in CellElements(row).ToList()) {
            long count = OdsRepeatModel.Read(element, OdfNamespaces.Table + "number-columns-repeated");
            if (columnIndex < checked(start + count)) {
                XElement target = OdsRepeatModel.Split(element, OdfNamespaces.Table + "number-columns-repeated", columnIndex - start);
                Dirty();
                return target;
            }
            start = checked(start + count);
        }
        long required = checked(columnIndex - start + 1);
        var added = new XElement(OdfNamespaces.Table + "table-cell");
        OdsRepeatModel.Set(added, OdfNamespaces.Table + "number-columns-repeated", required);
        row.Add(added);
        XElement result = OdsRepeatModel.Split(added, OdfNamespaces.Table + "number-columns-repeated", required - 1);
        Dirty();
        return result;
    }

    private IEnumerable<XElement> RowElements() => OdfTableRowElements.Enumerate(Element);
    internal static IEnumerable<XElement> CellElements(XElement row) => row.Elements()
        .Where(element => element.Name == OdfNamespaces.Table + "table-cell" || element.Name == OdfNamespaces.Table + "covered-table-cell");
    private void Dirty() => _document.MarkPartDirty("content.xml");
}
