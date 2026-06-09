namespace OfficeIMO.Pdf;

/// <summary>
/// Represents one normalized table extracted from a logical PDF page.
/// </summary>
public sealed class PdfLogicalTableExtraction {
    internal PdfLogicalTableExtraction(
        int pageIndex,
        int pageNumber,
        int tableIndex,
        PdfLogicalTable table,
        PdfLogicalTableData data) {
        PageIndex = pageIndex;
        PageNumber = pageNumber;
        TableIndex = tableIndex;
        Table = table;
        Data = data;
    }

    /// <summary>Zero-based page index in the extraction page collection.</summary>
    public int PageIndex { get; }

    /// <summary>One-based source page number from the PDF document.</summary>
    public int PageNumber { get; }

    /// <summary>Zero-based table index within the logical page.</summary>
    public int TableIndex { get; }

    /// <summary>Original logical table object with geometry and cell-level source data.</summary>
    public PdfLogicalTable Table { get; }

    /// <summary>Normalized table data suitable for readers and document-conversion adapters.</summary>
    public PdfLogicalTableData Data { get; }

    /// <summary>Detection heuristic that produced the source table.</summary>
    public string DetectionKind => Table.DetectionKind;
}
