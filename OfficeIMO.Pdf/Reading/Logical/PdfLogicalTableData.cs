namespace OfficeIMO.Pdf;

/// <summary>
/// Represents a structured, normalized view of a logical PDF table for downstream extraction surfaces.
/// </summary>
public sealed class PdfLogicalTableData {
    internal PdfLogicalTableData(
        PdfLogicalTableStructure structure,
        PdfLogicalTableDiagnostics diagnostics,
        IReadOnlyList<IReadOnlyList<string>> rows,
        IReadOnlyList<bool> numericColumns,
        bool truncated) {
        Structure = structure;
        Diagnostics = diagnostics;
        Columns = structure.Columns;
        Rows = rows;
        NumericColumns = SnapshotNumericColumns(numericColumns);
        ColumnProfiles = BuildColumnProfiles(Columns, Rows);
        TotalRowCount = structure.TotalBodyRowCount;
        Truncated = truncated;
    }

    /// <summary>Inferred schema and body-row boundaries used to build this table data.</summary>
    public PdfLogicalTableStructure Structure { get; }

    /// <summary>Confidence and geometry diagnostics for the detected source table.</summary>
    public PdfLogicalTableDiagnostics Diagnostics { get; }

    /// <summary>Inferred column names suitable for structured extraction surfaces.</summary>
    public IReadOnlyList<string> Columns { get; }

    /// <summary>Body rows padded or trimmed to the inferred column count.</summary>
    public IReadOnlyList<IReadOnlyList<string>> Rows { get; }

    /// <summary>Numeric body-column flags aligned to <see cref="Columns"/>.</summary>
    public IReadOnlyList<bool> NumericColumns { get; }

    /// <summary>Inferred column profiles aligned to <see cref="Columns"/>.</summary>
    public IReadOnlyList<PdfLogicalTableColumnProfile> ColumnProfiles { get; }

    /// <summary>Total body/data row count before any extraction cap was applied.</summary>
    public int TotalRowCount { get; }

    /// <summary>True when <see cref="Rows"/> contains fewer rows than <see cref="TotalRowCount"/>.</summary>
    public bool Truncated { get; }

    /// <summary>
    /// Reports whether the normalized table column has a numeric profile.
    /// </summary>
    /// <param name="columnIndex">Zero-based column index.</param>
    /// <returns>True when the column profile is numeric; otherwise false.</returns>
    public bool IsNumericColumn(int columnIndex) {
        return columnIndex >= 0 &&
            columnIndex < ColumnProfiles.Count &&
            ColumnProfiles[columnIndex].IsNumeric;
    }

    private static System.Collections.ObjectModel.ReadOnlyCollection<bool> SnapshotNumericColumns(IReadOnlyList<bool> numericColumns) {
        var copy = new bool[numericColumns.Count];
        for (int i = 0; i < numericColumns.Count; i++) {
            copy[i] = numericColumns[i];
        }

        return Array.AsReadOnly(copy);
    }

    private static System.Collections.ObjectModel.ReadOnlyCollection<PdfLogicalTableColumnProfile> BuildColumnProfiles(
        IReadOnlyList<string> columns,
        IReadOnlyList<IReadOnlyList<string>> rows) {
        var profiles = new PdfLogicalTableColumnProfile[columns.Count];
        for (int columnIndex = 0; columnIndex < columns.Count; columnIndex++) {
            int nonEmptyCount = 0;
            int numericCount = 0;
            for (int rowIndex = 0; rowIndex < rows.Count; rowIndex++) {
                IReadOnlyList<string> row = rows[rowIndex];
                string value = columnIndex < row.Count ? row[columnIndex] : string.Empty;
                if (string.IsNullOrWhiteSpace(value)) {
                    continue;
                }

                nonEmptyCount++;
                if (PdfLogicalTableAnalysis.LooksLikeNumericValue(value)) {
                    numericCount++;
                }
            }

            PdfLogicalTableColumnKind kind = GetColumnKind(nonEmptyCount, numericCount);
            profiles[columnIndex] = new PdfLogicalTableColumnProfile(
                columnIndex,
                columns[columnIndex],
                kind,
                nonEmptyCount,
                numericCount,
                GetConfidence(kind, nonEmptyCount, numericCount));
        }

        return Array.AsReadOnly(profiles);
    }

    private static PdfLogicalTableColumnKind GetColumnKind(int nonEmptyCount, int numericCount) {
        if (nonEmptyCount == 0) {
            return PdfLogicalTableColumnKind.Empty;
        }

        if (numericCount == nonEmptyCount) {
            return PdfLogicalTableColumnKind.Numeric;
        }

        return numericCount == 0 ? PdfLogicalTableColumnKind.Text : PdfLogicalTableColumnKind.Mixed;
    }

    private static double GetConfidence(PdfLogicalTableColumnKind kind, int nonEmptyCount, int numericCount) {
        if (kind == PdfLogicalTableColumnKind.Empty || nonEmptyCount == 0) {
            return 0d;
        }

        int matchingCount = kind == PdfLogicalTableColumnKind.Numeric
            ? numericCount
            : kind == PdfLogicalTableColumnKind.Text
                ? nonEmptyCount - numericCount
                : Math.Max(numericCount, nonEmptyCount - numericCount);
        return (double)matchingCount / nonEmptyCount;
    }
}
