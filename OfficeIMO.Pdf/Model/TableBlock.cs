namespace OfficeIMO.Pdf;

internal sealed class TableBlock : IPdfBlock {
    private readonly System.Collections.Generic.Dictionary<(int Row, int Col), string> _links = new();
    private readonly System.Collections.ObjectModel.ReadOnlyDictionary<(int Row, int Col), string> _linksView;

    public System.Collections.Generic.IReadOnlyList<string[]> Rows { get; }
    public System.Collections.Generic.IReadOnlyList<System.Collections.Generic.IReadOnlyList<PdfTableCell>> Cells { get; }
    public int ColumnCount { get; }
    public PdfAlign Align { get; }
    public PdfTableStyle? Style { get; }
    // Optional per-cell link URIs: key = (rowIndex, colIndex)
    public System.Collections.Generic.IReadOnlyDictionary<(int Row, int Col), string> Links => _linksView;

    public TableBlock(System.Collections.Generic.IEnumerable<string[]> rows, PdfAlign align, PdfTableStyle? style) {
        Guard.NotNull(rows, nameof(rows));
        Guard.LeftCenterRightAlign(align, nameof(align), "Table");
        _linksView = new System.Collections.ObjectModel.ReadOnlyDictionary<(int Row, int Col), string>(_links);
        Align = align; Style = style?.Clone();
        var cellSnapshot = new System.Collections.Generic.List<System.Collections.Generic.IReadOnlyList<PdfTableCell>>();
        foreach (var r in rows) {
            if (r is null) throw new System.ArgumentException("Table rows cannot contain null entries.", nameof(rows));
            var cells = new PdfTableCell[r.Length];
            for (int i = 0; i < r.Length; i++) {
                cells[i] = new PdfTableCell(r[i]);
            }

            cellSnapshot.Add(System.Array.AsReadOnly(cells));
        }

        Cells = cellSnapshot.AsReadOnly();
        ValidateMergedCellGrid(Cells);
        ValidateCellNamedDestinationNames(Cells);
        Rows = CreateTextRows(Cells);
        ColumnCount = GetColumnCount(Cells);
    }

    public TableBlock(System.Collections.Generic.IEnumerable<PdfTableCell[]> rows, PdfAlign align, PdfTableStyle? style) {
        Guard.NotNull(rows, nameof(rows));
        Guard.LeftCenterRightAlign(align, nameof(align), "Table");
        _linksView = new System.Collections.ObjectModel.ReadOnlyDictionary<(int Row, int Col), string>(_links);
        Align = align; Style = style?.Clone();
        var cellSnapshot = new System.Collections.Generic.List<System.Collections.Generic.IReadOnlyList<PdfTableCell>>();
        foreach (var r in rows) {
            if (r is null) throw new System.ArgumentException("Table rows cannot contain null entries.", nameof(rows));
            var cells = new PdfTableCell[r.Length];
            for (int i = 0; i < r.Length; i++) {
                if (r[i] is null) throw new System.ArgumentException("Table cells cannot contain null entries.", nameof(rows));
                cells[i] = r[i].Clone();
            }

            cellSnapshot.Add(System.Array.AsReadOnly(cells));
        }

        Cells = cellSnapshot.AsReadOnly();
        ValidateMergedCellGrid(Cells);
        ValidateCellNamedDestinationNames(Cells);
        Rows = CreateTextRows(Cells);
        ColumnCount = GetColumnCount(Cells);
    }

    internal void AddLink((int Row, int Col) cell, string uri) {
        _links[cell] = uri;
    }

    private static System.Collections.ObjectModel.ReadOnlyCollection<string[]> CreateTextRows(System.Collections.Generic.IReadOnlyList<System.Collections.Generic.IReadOnlyList<PdfTableCell>> rows) {
        var snapshot = new System.Collections.Generic.List<string[]>();
        for (int rowIndex = 0; rowIndex < rows.Count; rowIndex++) {
            var row = rows[rowIndex];
            var texts = new string[row.Count];
            for (int cellIndex = 0; cellIndex < row.Count; cellIndex++) {
                texts[cellIndex] = row[cellIndex].Text;
            }

            snapshot.Add(texts);
        }

        return snapshot.AsReadOnly();
    }

    private static void ValidateMergedCellGrid(System.Collections.Generic.IReadOnlyList<System.Collections.Generic.IReadOnlyList<PdfTableCell>> rows) {
        for (int rowIndex = 0; rowIndex < rows.Count; rowIndex++) {
            var row = rows[rowIndex];
            for (int cellIndex = 0; cellIndex < row.Count; cellIndex++) {
                PdfTableCell cell = row[cellIndex];
                if (cell.RowSpan > rows.Count - rowIndex) {
                    throw new System.ArgumentException("Table cell row span cannot extend beyond the available table rows.", nameof(rows));
                }
            }
        }
    }

    private static void ValidateCellNamedDestinationNames(System.Collections.Generic.IReadOnlyList<System.Collections.Generic.IReadOnlyList<PdfTableCell>> rows) {
        var names = new System.Collections.Generic.HashSet<string>(System.StringComparer.Ordinal);
        for (int rowIndex = 0; rowIndex < rows.Count; rowIndex++) {
            var row = rows[rowIndex];
            for (int cellIndex = 0; cellIndex < row.Count; cellIndex++) {
                string? name = row[cellIndex].NamedDestinationName;
                if (string.IsNullOrWhiteSpace(name)) {
                    continue;
                }

                if (!names.Add(name!)) {
                    throw new System.ArgumentException("Table cell named destinations must be unique.", nameof(rows));
                }
            }
        }
    }

    private static int GetColumnCount(System.Collections.Generic.IReadOnlyList<System.Collections.Generic.IReadOnlyList<PdfTableCell>> rows) {
        int columnCount = 0;
        var activeRowSpans = new System.Collections.Generic.List<int>();
        for (int rowIndex = 0; rowIndex < rows.Count; rowIndex++) {
            int column = 0;
            var row = rows[rowIndex];
            for (int cellIndex = 0; cellIndex < row.Count; cellIndex++) {
                while (column < activeRowSpans.Count && activeRowSpans[column] > 0) {
                    column++;
                }

                PdfTableCell cell = row[cellIndex];
                int lastColumn = column + cell.ColumnSpan;
                while (activeRowSpans.Count < lastColumn) {
                    activeRowSpans.Add(0);
                }

                for (int c = column; c < lastColumn; c++) {
                    activeRowSpans[c] = System.Math.Max(activeRowSpans[c], cell.RowSpan);
                }

                column = lastColumn;
            }

            columnCount = System.Math.Max(columnCount, activeRowSpans.Count);
            for (int c = 0; c < activeRowSpans.Count; c++) {
                if (activeRowSpans[c] > 0) {
                    activeRowSpans[c]--;
                }
            }
        }

        return columnCount;
    }
}
