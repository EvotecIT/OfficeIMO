namespace OfficeIMO.Pdf;

/// <summary>
/// Detected leader row such as a table-of-contents or label/value row.
/// </summary>
public sealed class PdfLogicalLeaderRow : IPdfLogicalElement {
    internal PdfLogicalLeaderRow(int pageNumber, string label, string value) {
        PageNumber = pageNumber;
        Label = label;
        Value = value;
    }

    /// <inheritdoc />
    public int PageNumber { get; }

    /// <inheritdoc />
    public PdfLogicalElementKind Kind => PdfLogicalElementKind.LeaderRow;

    /// <summary>Leader row label.</summary>
    public string Label { get; }

    /// <summary>Leader row trailing value.</summary>
    public string Value { get; }
}

/// <summary>
/// Detected table-like region with simple geometry.
/// </summary>
public sealed class PdfLogicalTable : IPdfLogicalElement {
    private PdfLogicalTable(
        int pageNumber,
        string kind,
        double yTop,
        double yBottom,
        IReadOnlyList<PdfLogicalTableColumn> columns,
        IReadOnlyList<IReadOnlyList<string>> rows,
        IReadOnlyList<PdfLogicalTableCell> cells) {
        PageNumber = pageNumber;
        DetectionKind = kind;
        YTop = yTop;
        YBottom = yBottom;
        Columns = columns;
        Rows = rows;
        Cells = cells;
        int expectedCells = rows.Count * columns.Count;
        int filledCells = rows.Sum(static row => row.Count(static cell => !string.IsNullOrWhiteSpace(cell)));
        double completeness = expectedCells == 0 ? 0D : (double)filledCells / expectedCells;
        Confidence = PdfInference.Clamp((columns.Count > 1 ? 0.45D : 0.2D) + (completeness * 0.45D) + (yTop > yBottom ? 0.1D : 0D));
        Evidence = new[] {
            new PdfInferenceEvidence("table.detection-kind", "The table was produced by the " + kind + " detector.", 0.5D),
            new PdfInferenceEvidence("table.cell-completeness", "Filled-cell completeness is " + completeness.ToString("0.###", System.Globalization.CultureInfo.InvariantCulture) + ".", (completeness * 2D) - 1D),
            new PdfInferenceEvidence("table.column-geometry", columns.Count > 1 ? "Multiple column boundaries were detected." : "Fewer than two column boundaries were detected.", columns.Count > 1 ? 0.7D : -0.5D)
        };
    }

    /// <inheritdoc />
    public int PageNumber { get; }

    /// <inheritdoc />
    public PdfLogicalElementKind Kind => PdfLogicalElementKind.Table;

    /// <summary>Detection heuristic that produced the table.</summary>
    public string DetectionKind { get; }

    /// <summary>Top Y coordinate of the detected table band.</summary>
    public double YTop { get; }

    /// <summary>Bottom Y coordinate of the detected table band.</summary>
    public double YBottom { get; }

    /// <summary>Detected table columns.</summary>
    public IReadOnlyList<PdfLogicalTableColumn> Columns { get; }

    /// <summary>Extracted table rows.</summary>
    public IReadOnlyList<IReadOnlyList<string>> Rows { get; }

    /// <summary>Extracted table cells with row and column indexes.</summary>
    public IReadOnlyList<PdfLogicalTableCell> Cells { get; }
    /// <summary>Normalized table-detection confidence.</summary>
    public double Confidence { get; }
    /// <summary>Evidence supporting the table detection.</summary>
    public IReadOnlyList<PdfInferenceEvidence> Evidence { get; }

    internal static PdfLogicalTable From(int pageNumber, StructuredTable table) {
        var columns = new List<PdfLogicalTableColumn>(table.Columns.Count);
        for (int i = 0; i < table.Columns.Count; i++) {
            columns.Add(new PdfLogicalTableColumn(table.Columns[i].From, table.Columns[i].To));
        }

        var rows = new List<IReadOnlyList<string>>(table.Rows.Count);
        var cells = new List<PdfLogicalTableCell>();
        for (int i = 0; i < table.Rows.Count; i++) {
            string[] row = (string[])table.Rows[i].Clone();
            rows.Add(Array.AsReadOnly(row));
            for (int columnIndex = 0; columnIndex < row.Length; columnIndex++) {
                PdfLogicalTableColumn? column = columnIndex < columns.Count ? columns[columnIndex] : null;
                cells.Add(new PdfLogicalTableCell(pageNumber, i, columnIndex, row[columnIndex], column));
            }
        }

        return new PdfLogicalTable(
            pageNumber,
            table.Kind,
            table.YTop,
            table.YBottom,
            columns.AsReadOnly(),
            rows.AsReadOnly(),
            cells.AsReadOnly());
    }
}

/// <summary>
/// Extracted table cell with row and column indexes.
/// </summary>
public sealed class PdfLogicalTableCell {
    internal PdfLogicalTableCell(int pageNumber, int rowIndex, int columnIndex, string text, PdfLogicalTableColumn? column) {
        PageNumber = pageNumber;
        RowIndex = rowIndex;
        ColumnIndex = columnIndex;
        Text = text;
        Column = column;
    }

    /// <summary>One-based source page number.</summary>
    public int PageNumber { get; }

    /// <summary>Zero-based row index within the detected table.</summary>
    public int RowIndex { get; }

    /// <summary>Zero-based column index within the detected table row.</summary>
    public int ColumnIndex { get; }

    /// <summary>Extracted cell text.</summary>
    public string Text { get; }

    /// <summary>Detected column geometry when available.</summary>
    public PdfLogicalTableColumn? Column { get; }
}

/// <summary>
/// Detected table column geometry.
/// </summary>
public sealed class PdfLogicalTableColumn {
    internal PdfLogicalTableColumn(double from, double to) {
        From = from;
        To = to;
    }

    /// <summary>Left X coordinate in PDF points.</summary>
    public double From { get; }

    /// <summary>Right X coordinate in PDF points.</summary>
    public double To { get; }
}
