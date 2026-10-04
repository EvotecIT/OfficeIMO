namespace OfficeIMO.OpenDocument;

/// <summary>An XML-backed ODS worksheet with sparse repeat-run editing.</summary>
public sealed partial class OdsSheet {
    /// <summary>Default maximum number of cells that one merge operation may materialize.</summary>
    public const long DefaultMaximumMergeCells = 100_000;
    /// <summary>Maximum number of embedded chart frames projected from one sheet.</summary>
    public const int DefaultMaximumChartFrames = 256;

    private readonly OdsDocument _document;
    private int _editExternalVersion = -1;
    private XElement? _lastEditedRow;
    private long _lastEditedRowStart;
    private XElement? _lastEditedCellRow;
    private XElement? _lastEditedCell;
    private long _lastEditedCellStart;
    private IReadOnlyList<OdsColumnRun>? _cachedColumnRuns;
    private int _columnRunsContentVersion = -1;
    private int _columnRunsExternalVersion = -1;
    private bool _cachedHasColumnDefaults;

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

    /// <summary>Embedded charts whose chart content can be read safely from this package.</summary>
    public IReadOnlyList<OdsChart> Charts {
        get {
            IReadOnlyList<OdsChart> charts = GetCharts(DefaultMaximumChartFrames, out bool truncated);
            if (truncated) throw new NotSupportedException("The sheet exceeds the embedded chart frame limit.");
            return charts;
        }
    }

    internal IReadOnlyList<OdsChart> GetCharts(int maximumFrames, out bool truncated) {
        if (maximumFrames < 1) throw new ArgumentOutOfRangeException(nameof(maximumFrames));
        var charts = new List<OdsChart>();
        var parsedParts = new Dictionary<string, OdsChart?>(StringComparer.Ordinal);
        int frames = 0;
        bool reachedLimit = false;
        void AddFrame(XElement frame, long? anchorRow, long? anchorColumn) {
            if (frame.Element(OdfNamespaces.Draw + "object") == null) return;
            if (++frames > maximumFrames) { reachedLimit = true; return; }
            if (!OdsChart.TryGetContentPath(frame, out _, out string partPath)) return;
            if (!parsedParts.TryGetValue(partPath, out OdsChart? template)) {
                template = OdsChart.TryRead(_document, frame, anchorRow, anchorColumn);
                parsedParts.Add(partPath, template);
                if (template != null) charts.Add(template);
            } else if (template != null) {
                try { charts.Add(template.WithFrame(frame, anchorRow, anchorColumn)); }
                catch (InvalidDataException) { /* A malformed frame is not a readable chart. */ }
            }
        }
        long rowIndex = 0;
        foreach (XElement row in RowElements()) {
            long rowRepeat = OdsRepeatModel.Read(row, OdfNamespaces.Table + "number-rows-repeated");
            long columnIndex = 0;
            foreach (XElement cell in CellElements(row)) {
                long columnRepeat = OdsRepeatModel.Read(cell, OdfNamespaces.Table + "number-columns-repeated");
                if (rowRepeat == 1 && columnRepeat == 1) {
                    foreach (XElement frame in cell.Descendants(OdfNamespaces.Draw + "frame")) {
                        AddFrame(frame, rowIndex, columnIndex);
                        if (reachedLimit) break;
                    }
                }
                if (reachedLimit) break;
                columnIndex = checked(columnIndex + columnRepeat);
            }
            if (reachedLimit) break;
            rowIndex = checked(rowIndex + rowRepeat);
        }
        XElement? shapes = Element.Element(OdfNamespaces.Table + "shapes");
        if (!reachedLimit && shapes != null) {
            foreach (XElement frame in shapes.Descendants(OdfNamespaces.Draw + "frame")) {
                AddFrame(frame, null, null);
                if (reachedLimit) break;
            }
        }
        truncated = reachedLimit;
        return charts;
    }

    /// <summary>Optional ODF print range expression.</summary>
    public string? PrintRanges {
        get => (string?)Element.Attribute(OdfNamespaces.Table + "print-ranges");
        set { Element.SetAttributeValue(OdfNamespaces.Table + "print-ranges", value); Dirty(); }
    }

    /// <summary>Sparse row runs without expanding <c>table:number-rows-repeated</c>.</summary>
    public IReadOnlyList<OdsRowRun> RowRuns => GetRowRuns();

    internal IReadOnlyList<OdsRowRun> GetRowRuns(IReadOnlyList<OdsColumnRun>? columnRuns = null) {
        var runs = new List<OdsRowRun>();
        Func<IReadOnlyList<OdsColumnRun>> getColumns;
        Func<bool> hasColumnDefaults;
        if (columnRuns == null) {
            getColumns = GetCachedColumnRuns;
            hasColumnDefaults = () => _cachedHasColumnDefaults;
        } else {
            IReadOnlyList<OdsColumnRun> fixedColumns = columnRuns;
            bool fixedHasDefaults = fixedColumns.Any(column => column.DefaultCellStyleName != null);
            getColumns = () => fixedColumns;
            hasColumnDefaults = () => fixedHasDefaults;
        }
        long start = 0;
        foreach (XElement row in RowElements()) {
            long repeat = OdsRepeatModel.Read(row, OdfNamespaces.Table + "number-rows-repeated");
            runs.Add(new OdsRowRun(_document, row, start, repeat,
                column => GetDefaultCellStyleName(row, column, getColumns()),
                getColumns, hasColumnDefaults));
            start = checked(start + repeat);
        }
        return runs;
    }

    private IReadOnlyList<OdsColumnRun> GetCachedColumnRuns() {
        if (_cachedColumnRuns == null || _columnRunsContentVersion != _document.Package.ContentEditVersion ||
            _columnRunsExternalVersion != _document.Package.ExternalXmlEditVersion) {
            _cachedColumnRuns = ColumnRuns;
            _cachedHasColumnDefaults = _cachedColumnRuns.Any(column => column.DefaultCellStyleName != null);
            _columnRunsContentVersion = _document.Package.ContentEditVersion;
            _columnRunsExternalVersion = _document.Package.ExternalXmlEditVersion;
        }
        return _cachedColumnRuns;
    }

    /// <summary>Sparse column definition runs without expanding repeats.</summary>
    public IReadOnlyList<OdsColumnRun> ColumnRuns {
        get {
            var runs = new List<OdsColumnRun>();
            long start = 0;
            foreach (XElement column in ColumnElements()) {
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
        return new OdsCell(_document, cellElement,
            inheritedStyleResolver: () => GetDefaultCellStyleName(rowElement, column));
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
        foreach (XElement element in ColumnElements().ToList()) {
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
        XElement? insertionPoint = Element.Elements().FirstOrDefault(child => child.Name == OdfNamespaces.Table + "table-row"
            || child.Name == OdfNamespaces.Table + "table-header-rows"
            || child.Name == OdfNamespaces.Table + "table-rows"
            || child.Name == OdfNamespaces.Table + "table-row-group");
        if (insertionPoint == null) Element.Add(added); else insertionPoint.AddBeforeSelf(added);
        XElement result = OdsRepeatModel.Split(added, OdfNamespaces.Table + "number-columns-repeated", required - 1);
        Dirty();
        return new OdsColumn(_document, result);
    }

    /// <summary>Reads a value without splitting or expanding repeat runs.</summary>
    public OdsCellValue GetValue(long row, long column) {
        if (row < 0) throw new ArgumentOutOfRangeException(nameof(row));
        if (column < 0) throw new ArgumentOutOfRangeException(nameof(column));
        XElement? cell = FindPrototypeCell(row, column);
        return cell == null ? OdsCellValue.Empty : OdsCell.ReadValue(cell);
    }

    /// <summary>Reads a formula without splitting or expanding repeat runs.</summary>
    public string? GetFormula(long row, long column) {
        if (row < 0) throw new ArgumentOutOfRangeException(nameof(row));
        if (column < 0) throw new ArgumentOutOfRangeException(nameof(column));
        XElement? cell = FindPrototypeCell(row, column);
        return (string?)cell?.Attribute(OdfNamespaces.Table + "formula");
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

    internal static string? GetDefaultCellStyleName(OdsRowRun row, long column, IReadOnlyList<OdsColumnRun> columns) =>
        GetDefaultCellStyleName(row.Element, column, columns);

    private string? GetDefaultCellStyleName(XElement row, long column) =>
        GetDefaultCellStyleName(row, column, ColumnRuns);

    private static string? GetDefaultCellStyleName(XElement row, long column, IReadOnlyList<OdsColumnRun> columns) {
        string? rowStyle = (string?)row.Attribute(OdfNamespaces.Table + "default-cell-style-name");
        if (rowStyle != null) return rowStyle;
        int low = 0, high = columns.Count - 1;
        while (low <= high) {
            int middle = low + (high - low) / 2;
            OdsColumnRun definition = columns[middle];
            if (column < definition.StartColumn) high = middle - 1;
            else if (column >= checked(definition.StartColumn + definition.RepeatCount)) low = middle + 1;
            else return definition.DefaultCellStyleName;
        }
        return null;
    }

    private IEnumerable<XElement> ColumnElements() => Element.Descendants(OdfNamespaces.Table + "table-column")
        .Where(column => ReferenceEquals(column.Ancestors(OdfNamespaces.Table + "table").FirstOrDefault(), Element));

    private XElement GetRowForEdit(long rowIndex) {
        if (_editExternalVersion != _document.Package.ExternalXmlEditVersion) {
            _lastEditedRow = null;
            _lastEditedCellRow = null;
            _lastEditedCell = null;
            _editExternalVersion = _document.Package.ExternalXmlEditVersion;
        }
        long start = 0;
        IEnumerable<XElement> candidates = RowElements();
        if (_lastEditedRow?.Parent != null && rowIndex >= _lastEditedRowStart) {
            long cachedCount = OdsRepeatModel.Read(_lastEditedRow, OdfNamespaces.Table + "number-rows-repeated");
            if (rowIndex < checked(_lastEditedRowStart + cachedCount)) {
                XElement cachedTarget = OdsRepeatModel.Split(
                    _lastEditedRow,
                    OdfNamespaces.Table + "number-rows-repeated",
                    rowIndex - _lastEditedRowStart);
                CacheRow(cachedTarget, rowIndex);
                Dirty();
                return cachedTarget;
            }
            start = checked(_lastEditedRowStart + cachedCount);
            candidates = OdfTableRowElements.EnumerateAfter(Element, _lastEditedRow);
        }

        foreach (XElement element in candidates) {
            long count = OdsRepeatModel.Read(element, OdfNamespaces.Table + "number-rows-repeated");
            if (rowIndex < checked(start + count)) {
                XElement target = OdsRepeatModel.Split(element, OdfNamespaces.Table + "number-rows-repeated", rowIndex - start);
                CacheRow(target, rowIndex);
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
        CacheRow(result, rowIndex);
        Dirty();
        return result;
    }

    private XElement GetCellForEdit(XElement row, long columnIndex) {
        long start = 0;
        IEnumerable<XElement> candidates = CellElements(row);
        if (ReferenceEquals(_lastEditedCellRow, row)
            && _lastEditedCell?.Parent != null
            && columnIndex >= _lastEditedCellStart) {
            long cachedCount = OdsRepeatModel.Read(_lastEditedCell, OdfNamespaces.Table + "number-columns-repeated");
            if (columnIndex < checked(_lastEditedCellStart + cachedCount)) {
                XElement cachedTarget = OdsRepeatModel.Split(
                    _lastEditedCell,
                    OdfNamespaces.Table + "number-columns-repeated",
                    columnIndex - _lastEditedCellStart);
                CacheCell(row, cachedTarget, columnIndex);
                Dirty();
                return cachedTarget;
            }
            start = checked(_lastEditedCellStart + cachedCount);
            candidates = _lastEditedCell.ElementsAfterSelf()
                .Where(IsCellElement);
        }

        foreach (XElement element in candidates) {
            long count = OdsRepeatModel.Read(element, OdfNamespaces.Table + "number-columns-repeated");
            if (columnIndex < checked(start + count)) {
                XElement target = OdsRepeatModel.Split(element, OdfNamespaces.Table + "number-columns-repeated", columnIndex - start);
                CacheCell(row, target, columnIndex);
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
        CacheCell(row, result, columnIndex);
        Dirty();
        return result;
    }

    private void CacheRow(XElement row, long rowIndex) {
        _lastEditedRow = row;
        _lastEditedRowStart = rowIndex;
        if (!ReferenceEquals(_lastEditedCellRow, row)) {
            _lastEditedCellRow = null;
            _lastEditedCell = null;
            _lastEditedCellStart = 0;
        }
    }

    private void CacheCell(XElement row, XElement cell, long columnIndex) {
        _lastEditedCellRow = row;
        _lastEditedCell = cell;
        _lastEditedCellStart = columnIndex;
    }

    private IEnumerable<XElement> RowElements() => OdfTableRowElements.Enumerate(Element);
    internal static IEnumerable<XElement> CellElements(XElement row) => row.Elements().Where(IsCellElement);
    private static bool IsCellElement(XElement element) =>
        element.Name == OdfNamespaces.Table + "table-cell" || element.Name == OdfNamespaces.Table + "covered-table-cell";
    private void Dirty() => _document.MarkPartDirty("content.xml");
}
