namespace OfficeIMO.Pdf;

/// <summary>Describes a table added to the document flow. PDF tables are fixed once added; change the rows or style before adding.</summary>
public sealed class PdfTable {
    internal PdfTable(int rowCount, int columnCount, string? caption, bool isDeferred) {
        RowCount = rowCount; ColumnCount = columnCount; Caption = caption; IsDeferred = isDeferred;
    }

    /// <summary>Number of rows, including header rows. Zero for a deferred table, whose rows are read during layout.</summary>
    public int RowCount { get; }

    /// <summary>Number of logical columns. Zero for a deferred table, whose rows are read during layout.</summary>
    public int ColumnCount { get; }

    /// <summary>Visible caption text, if any.</summary>
    public string? Caption { get; }

    /// <summary>True when rows come from a row factory that is enumerated in batches during layout.</summary>
    public bool IsDeferred { get; }
}
