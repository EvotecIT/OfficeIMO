namespace OfficeIMO.Pdf;

/// <summary>Schema states that can be established from logical PDF table evidence.</summary>
public enum PdfLogicalTableSchemaKind {
    /// <summary>No row can be promoted to schema without guessing.</summary>
    Unknown,
    /// <summary>The first source row is supported as a column-header row.</summary>
    HeaderRow
}

/// <summary>
/// Describes the established schema and body-row boundaries for a logical PDF table.
/// </summary>
public sealed class PdfLogicalTableStructure {
    internal PdfLogicalTableStructure(
        int columnCount,
        IReadOnlyList<string> columns,
        int bodyStartRowIndex,
        int totalBodyRowCount,
        bool hasHeaderRow,
        PdfLogicalTableSchemaKind schemaKind,
        double schemaConfidence,
        IReadOnlyList<PdfInferenceEvidence> schemaEvidence) {
        ColumnCount = columnCount;
        Columns = SnapshotColumns(columns);
        BodyStartRowIndex = bodyStartRowIndex;
        TotalBodyRowCount = totalBodyRowCount;
        HasHeaderRow = hasHeaderRow;
        SchemaKind = schemaKind;
        SchemaConfidence = PdfInference.Clamp(schemaConfidence);
        SchemaEvidence = PdfInference.Snapshot(schemaEvidence);
    }

    /// <summary>Maximum visible cell count across table rows.</summary>
    public int ColumnCount { get; }

    /// <summary>Structurally established column names, or empty names when the schema is unknown.</summary>
    public IReadOnlyList<string> Columns { get; }

    /// <summary>Zero-based row index where body/data rows begin.</summary>
    public int BodyStartRowIndex { get; }

    /// <summary>Total body/data row count before any consumer-side truncation.</summary>
    public int TotalBodyRowCount { get; }

    /// <summary>True when the first logical row was promoted to column headers.</summary>
    public bool HasHeaderRow { get; }

    /// <summary>Strongest schema state supported by the available structural evidence.</summary>
    public PdfLogicalTableSchemaKind SchemaKind { get; }

    /// <summary>Normalized confidence in <see cref="SchemaKind"/> from 0 to 1.</summary>
    public double SchemaConfidence { get; }

    /// <summary>Evidence supporting the reported schema state.</summary>
    public IReadOnlyList<PdfInferenceEvidence> SchemaEvidence { get; }

    private static System.Collections.ObjectModel.ReadOnlyCollection<string> SnapshotColumns(IReadOnlyList<string> columns) {
        var copy = new string[columns.Count];
        for (int i = 0; i < columns.Count; i++) {
            copy[i] = columns[i] ?? string.Empty;
        }

        return Array.AsReadOnly(copy);
    }
}
