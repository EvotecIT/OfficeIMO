namespace OfficeIMO.Pdf;

/// <summary>
/// Diagnostic confidence and geometry signals for a normalized logical PDF table.
/// </summary>
public sealed class PdfLogicalTableDiagnostics {
    internal PdfLogicalTableDiagnostics(
        string detectionKind,
        double confidence,
        double schemaConfidence,
        double cellCompleteness,
        double columnGeometryConfidence,
        int sourceRowCount,
        int expectedCellCount,
        int filledCellCount,
        double xStart,
        double xEnd,
        double yTop,
        double yBottom,
        bool hasGeometry) {
        DetectionKind = detectionKind ?? string.Empty;
        Confidence = confidence;
        SchemaConfidence = schemaConfidence;
        CellCompleteness = cellCompleteness;
        ColumnGeometryConfidence = columnGeometryConfidence;
        SourceRowCount = sourceRowCount;
        ExpectedCellCount = expectedCellCount;
        FilledCellCount = filledCellCount;
        MissingCellCount = Math.Max(0, expectedCellCount - filledCellCount);
        XStart = xStart;
        XEnd = xEnd;
        YTop = yTop;
        YBottom = yBottom;
        HasGeometry = hasGeometry;
        Evidence = BuildEvidence(DetectionKind, schemaConfidence, cellCompleteness, columnGeometryConfidence);
    }

    /// <summary>Detection heuristic that produced the source logical table.</summary>
    public string DetectionKind { get; }

    /// <summary>Overall confidence score between 0 and 1 based on schema, cell completeness, and column geometry signals.</summary>
    public double Confidence { get; }

    /// <summary>Confidence score between 0 and 1 for inferred table schema, including header and key/value recognition.</summary>
    public double SchemaConfidence { get; }

    /// <summary>Ratio between 0 and 1 of non-empty cells to expected cells in the detected source table.</summary>
    public double CellCompleteness { get; }

    /// <summary>Confidence score between 0 and 1 for detected column geometry matching the inferred table width.</summary>
    public double ColumnGeometryConfidence { get; }

    /// <summary>Number of source rows detected before body-row normalization.</summary>
    public int SourceRowCount { get; }

    /// <summary>Expected source cell count from source rows multiplied by inferred columns.</summary>
    public int ExpectedCellCount { get; }

    /// <summary>Number of non-empty source cells detected.</summary>
    public int FilledCellCount { get; }

    /// <summary>Number of expected source cells that were empty or unavailable.</summary>
    public int MissingCellCount { get; }

    /// <summary>Left edge of the detected table geometry in PDF points.</summary>
    public double XStart { get; }

    /// <summary>Right edge of the detected table geometry in PDF points.</summary>
    public double XEnd { get; }

    /// <summary>Top baseline coordinate of the detected table geometry in PDF points.</summary>
    public double YTop { get; }

    /// <summary>Bottom baseline coordinate of the detected table geometry in PDF points.</summary>
    public double YBottom { get; }

    /// <summary>Detected table width in PDF points.</summary>
    public double Width => Math.Max(0D, XEnd - XStart);

    /// <summary>Detected table height in PDF points.</summary>
    public double Height => Math.Max(0D, YTop - YBottom);

    /// <summary>True when table and column coordinates were available.</summary>
    public bool HasGeometry { get; }
    /// <summary>Stable diagnostic evidence behind the component confidence scores.</summary>
    public IReadOnlyList<PdfInferenceEvidence> Evidence { get; }

    internal static PdfLogicalTableDiagnostics Create(PdfLogicalTable table, PdfLogicalTableStructure structure) {
        int sourceRowCount = table.Rows.Count;
        int columnCount = Math.Max(0, structure.ColumnCount);
        int expectedCellCount = sourceRowCount * columnCount;
        int filledCellCount = CountFilledCells(table.Rows, columnCount);
        double cellCompleteness = expectedCellCount == 0
            ? 0D
            : Clamp01((double)filledCellCount / expectedCellCount);
        double schemaConfidence = GetSchemaConfidence(structure);
        double columnGeometryConfidence = GetColumnGeometryConfidence(table, columnCount);
        double confidence = Clamp01(
            (schemaConfidence * 0.25D) +
            (cellCompleteness * 0.35D) +
            (columnGeometryConfidence * 0.40D));

        (double xStart, double xEnd, bool hasGeometry) = GetHorizontalGeometry(table);
        return new PdfLogicalTableDiagnostics(
            table.DetectionKind,
            confidence,
            schemaConfidence,
            cellCompleteness,
            columnGeometryConfidence,
            sourceRowCount,
            expectedCellCount,
            filledCellCount,
            xStart,
            xEnd,
            table.YTop,
            table.YBottom,
            hasGeometry);
    }

    private static int CountFilledCells(IReadOnlyList<IReadOnlyList<string>> rows, int columnCount) {
        int count = 0;
        for (int rowIndex = 0; rowIndex < rows.Count; rowIndex++) {
            IReadOnlyList<string> row = rows[rowIndex];
            int cellsToInspect = columnCount == 0 ? row.Count : Math.Min(columnCount, row.Count);
            for (int columnIndex = 0; columnIndex < cellsToInspect; columnIndex++) {
                if (!string.IsNullOrWhiteSpace(row[columnIndex])) {
                    count++;
                }
            }
        }

        return count;
    }

    private static double GetSchemaConfidence(PdfLogicalTableStructure structure) {
        if (structure.HasHeaderRow) {
            return 1D;
        }

        if (structure.IsKeyValueTable) {
            return 0.85D;
        }

        return structure.ColumnCount > 1 && structure.TotalBodyRowCount > 1 ? 0.65D : 0.4D;
    }

    private static double GetColumnGeometryConfidence(PdfLogicalTable table, int columnCount) {
        if (columnCount == 0 || table.Columns.Count == 0) {
            return 0D;
        }

        int comparableColumns = Math.Min(columnCount, table.Columns.Count);
        int positiveWidthColumns = 0;
        for (int i = 0; i < comparableColumns; i++) {
            if (table.Columns[i].To > table.Columns[i].From) {
                positiveWidthColumns++;
            }
        }

        double countScore = Clamp01((double)comparableColumns / columnCount);
        double widthScore = Clamp01((double)positiveWidthColumns / columnCount);
        double verticalScore = table.YTop > table.YBottom ? 1D : 0.5D;
        return Clamp01((countScore + widthScore + verticalScore) / 3D);
    }

    private static (double XStart, double XEnd, bool HasGeometry) GetHorizontalGeometry(PdfLogicalTable table) {
        if (table.Columns.Count == 0) {
            return (0D, 0D, false);
        }

        double xStart = double.MaxValue;
        double xEnd = double.MinValue;
        for (int i = 0; i < table.Columns.Count; i++) {
            xStart = Math.Min(xStart, table.Columns[i].From);
            xEnd = Math.Max(xEnd, table.Columns[i].To);
        }

        return (xStart, xEnd, xEnd > xStart);
    }

    private static double Clamp01(double value) {
        if (value <= 0D) {
            return 0D;
        }

        return value >= 1D ? 1D : value;
    }

    private static PdfInferenceEvidence[] BuildEvidence(string detectionKind, double schemaConfidence, double cellCompleteness, double geometryConfidence) => new[] {
        new PdfInferenceEvidence("table.detection-kind", "The table was produced by the " + detectionKind + " detector.", 0.5D),
        new PdfInferenceEvidence("table.schema-confidence", "Schema confidence is " + schemaConfidence.ToString("0.###", System.Globalization.CultureInfo.InvariantCulture) + ".", (schemaConfidence * 2D) - 1D),
        new PdfInferenceEvidence("table.cell-completeness", "Cell completeness is " + cellCompleteness.ToString("0.###", System.Globalization.CultureInfo.InvariantCulture) + ".", (cellCompleteness * 2D) - 1D),
        new PdfInferenceEvidence("table.geometry-confidence", "Column geometry confidence is " + geometryConfidence.ToString("0.###", System.Globalization.CultureInfo.InvariantCulture) + ".", (geometryConfidence * 2D) - 1D)
    };
}
