namespace OfficeIMO.Pdf;

/// <summary>
/// Describes the inferred schema and body-row boundaries for a logical PDF table.
/// </summary>
public sealed class PdfLogicalTableStructure {
    internal PdfLogicalTableStructure(
        int columnCount,
        IReadOnlyList<string> columns,
        int bodyStartRowIndex,
        int totalBodyRowCount,
        bool hasHeaderRow,
        bool isKeyValueTable) {
        ColumnCount = columnCount;
        Columns = SnapshotColumns(columns);
        BodyStartRowIndex = bodyStartRowIndex;
        TotalBodyRowCount = totalBodyRowCount;
        HasHeaderRow = hasHeaderRow;
        IsKeyValueTable = isKeyValueTable;
    }

    /// <summary>Maximum visible cell count across table rows.</summary>
    public int ColumnCount { get; }

    /// <summary>Inferred column names suitable for structured extraction surfaces.</summary>
    public IReadOnlyList<string> Columns { get; }

    /// <summary>Zero-based row index where body/data rows begin.</summary>
    public int BodyStartRowIndex { get; }

    /// <summary>Total body/data row count before any consumer-side truncation.</summary>
    public int TotalBodyRowCount { get; }

    /// <summary>True when the first logical row was promoted to column headers.</summary>
    public bool HasHeaderRow { get; }

    /// <summary>True when the table looks like a two-column key/value fact table.</summary>
    public bool IsKeyValueTable { get; }

    private static System.Collections.ObjectModel.ReadOnlyCollection<string> SnapshotColumns(IReadOnlyList<string> columns) {
        var copy = new string[columns.Count];
        for (int i = 0; i < columns.Count; i++) {
            copy[i] = columns[i] ?? string.Empty;
        }

        return Array.AsReadOnly(copy);
    }
}
