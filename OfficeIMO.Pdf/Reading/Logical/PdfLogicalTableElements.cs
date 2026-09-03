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
        IReadOnlyList<PdfLogicalTableCell> cells,
        PdfLogicalContentSourceKind sourceKind = PdfLogicalContentSourceKind.Native,
        PdfTableCoordinateSpace coordinateSpace = PdfTableCoordinateSpace.PdfUserSpace,
        PdfLogicalVisualBounds? visualBounds = null,
        double? confidence = null,
        IReadOnlyList<PdfInferenceEvidence>? evidence = null) {
        PageNumber = pageNumber;
        DetectionKind = kind;
        YTop = yTop;
        YBottom = yBottom;
        Columns = columns;
        Rows = rows;
        Cells = cells;
        SourceKind = sourceKind;
        CoordinateSpace = coordinateSpace;
        VisualBounds = visualBounds;
        int expectedCells = rows.Count * columns.Count;
        int filledCells = rows.Sum(static row => row.Count(static cell => !string.IsNullOrWhiteSpace(cell)));
        double completeness = expectedCells == 0 ? 0D : (double)filledCells / expectedCells;
        Confidence = PdfInference.Clamp(confidence ?? ((columns.Count > 1 ? 0.45D : 0.2D) + (completeness * 0.45D) + (visualBounds is not null || yTop > yBottom ? 0.1D : 0D)));
        Evidence = evidence ?? new[] {
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

    /// <summary>Top Y coordinate in <see cref="CoordinateSpace"/>.</summary>
    public double YTop { get; }

    /// <summary>Bottom Y coordinate in <see cref="CoordinateSpace"/>.</summary>
    public double YBottom { get; }

    /// <summary>Detected table columns in <see cref="CoordinateSpace"/>.</summary>
    public IReadOnlyList<PdfLogicalTableColumn> Columns { get; }

    /// <summary>Extracted table rows.</summary>
    public IReadOnlyList<IReadOnlyList<string>> Rows { get; }

    /// <summary>Extracted table cells with row and column indexes.</summary>
    public IReadOnlyList<PdfLogicalTableCell> Cells { get; }
    /// <summary>Whether table evidence came from native PDF operations or accepted OCR geometry.</summary>
    public PdfLogicalContentSourceKind SourceKind { get; }
    /// <summary>Coordinate system used by the table bounds and columns.</summary>
    public PdfTableCoordinateSpace CoordinateSpace { get; }
    /// <summary>Direct top-left visual geometry when the detector operates on rendered OCR coordinates.</summary>
    public PdfLogicalVisualBounds? VisualBounds { get; }
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

    internal static PdfLogicalTable From(int pageNumber, PdfUnderstandingTableCandidate table) {
        var columns = table.Columns
            .Select(static column => new PdfLogicalTableColumn(column.From, column.To))
            .ToArray();
        var rows = table.Rows
            .Select(static row => (IReadOnlyList<string>)Array.AsReadOnly(row.ToArray()))
            .ToArray();
        int cellCount = 0;
        for (int rowIndex = 0; rowIndex < rows.Length; rowIndex++) {
            cellCount = checked(cellCount + rows[rowIndex].Count);
        }
        var cells = new List<PdfLogicalTableCell>(cellCount);
        for (int rowIndex = 0; rowIndex < rows.Length; rowIndex++) {
            for (int columnIndex = 0; columnIndex < rows[rowIndex].Count; columnIndex++) {
                PdfLogicalTableColumn? column = columnIndex < columns.Length ? columns[columnIndex] : null;
                cells.Add(new PdfLogicalTableCell(pageNumber, rowIndex, columnIndex, rows[rowIndex][columnIndex], column));
            }
        }
        return new PdfLogicalTable(
            pageNumber,
            table.DetectionKind,
            table.YTop,
            table.YBottom,
            columns,
            rows,
            cells.AsReadOnly(),
            table.SourceKind,
            table.CoordinateSpace,
            table.VisualBounds,
            table.Confidence,
            table.Evidence);
    }

    internal static PdfLogicalTable FromOcr(
        int pageNumber,
        double top,
        double bottom,
        IReadOnlyList<(double From, double To)> columnBounds,
        IReadOnlyList<IReadOnlyList<string>> sourceRows) {
        double left = columnBounds.Min(static column => column.From);
        double right = columnBounds.Max(static column => column.To);
        PdfUnderstandingTableCandidate candidate = PdfUnderstandingTableCandidate.FromOcr(
            "OcrAlignedColumns",
            top,
            bottom,
            new PdfLogicalVisualBounds(left, top, right, bottom),
            columnBounds,
            sourceRows,
            0.8D,
            new[] {
                new PdfInferenceEvidence(
                    "table.ocr-aligned-geometry",
                    "Accepted OCR words form repeated aligned columns.",
                    0.8D)
            });
        return From(pageNumber, candidate);
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

    /// <summary>Left X coordinate in the owning table's coordinate space.</summary>
    public double From { get; }

    /// <summary>Right X coordinate in the owning table's coordinate space.</summary>
    public double To { get; }
}
