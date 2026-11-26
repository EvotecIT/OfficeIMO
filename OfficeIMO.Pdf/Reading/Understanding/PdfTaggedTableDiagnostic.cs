namespace OfficeIMO.Pdf;

/// <summary>Why tagged reconstruction was skipped for a PDF table; the page is then read by geometric detection, which may or may not find a table.</summary>
public enum PdfTaggedTableFallbackReason {
    /// <summary>The Table element holds children other than row groups, rows and cells, overlapping spans, or broken parent links.</summary>
    UnsupportedStructure,
    /// <summary>A row's cells carry content from more than one page, or rows go back to an earlier page.</summary>
    RowSpansPages,
    /// <summary>A cell refers to marked content that its page does not contain as text.</summary>
    UnresolvedContent,
    /// <summary>The tagged table holds no text.</summary>
    NoContent,
    /// <summary>Fewer than two rows hold text, or no row has two filled cells, so the table is page layout rather than data.</summary>
    LayoutOnly,
    /// <summary>Rows list their cells in an order that contradicts where the cells sit on the page.</summary>
    CellOrderConflictsWithLayout
}

/// <summary>Reports a tagged PDF table whose structure was not used to rebuild its grid.</summary>
public sealed class PdfTaggedTableDiagnostic {
    internal PdfTaggedTableDiagnostic(int structureObjectNumber, int pageNumber, PdfTaggedTableFallbackReason reason, string message) {
        StructureObjectNumber = structureObjectNumber;
        PageNumber = pageNumber;
        Reason = reason;
        Message = message;
    }

    /// <summary>Object number of the Table structure element.</summary>
    public int StructureObjectNumber { get; }

    /// <summary>One-based page the table's content is on.</summary>
    public int PageNumber { get; }

    /// <summary>Why the table structure was not used.</summary>
    public PdfTaggedTableFallbackReason Reason { get; }

    /// <summary>Human-readable explanation.</summary>
    public string Message { get; }

    /// <inheritdoc />
    public override string ToString() => "Tagged table " + StructureObjectNumber.ToString(System.Globalization.CultureInfo.InvariantCulture) +
        " on page " + PageNumber.ToString(System.Globalization.CultureInfo.InvariantCulture) + ": " + Message;
}

/// <summary>One page's rows of a tagged table, and the table they belong to.</summary>
internal sealed class PdfTaggedTableSegment {
    internal PdfTaggedTableSegment(
        int tableObjectNumber,
        int pageNumber,
        int segmentIndex,
        int headerRowCount,
        IReadOnlyList<string>? columnNames,
        IReadOnlyList<PdfTaggedTableCellSpan> spans) {
        TableObjectNumber = tableObjectNumber;
        PageNumber = pageNumber;
        SegmentIndex = segmentIndex;
        HeaderRowCount = headerRowCount;
        ColumnNames = columnNames;
        Spans = spans;
    }

    internal int TableObjectNumber { get; }
    internal int PageNumber { get; }
    /// <summary>Zero-based position of this page's rows among the table's pages.</summary>
    internal int SegmentIndex { get; }
    /// <summary>Number of tagged header rows this segment starts with; zero when it continues the table without them.</summary>
    internal int HeaderRowCount { get; }
    /// <summary>Column names from the table's tagged header rows, joined top to bottom, or null without a header.</summary>
    internal IReadOnlyList<string>? ColumnNames { get; }
    internal IReadOnlyList<PdfTaggedTableCellSpan> Spans { get; }
}

/// <summary>A merged cell in segment row and column coordinates.</summary>
internal readonly record struct PdfTaggedTableCellSpan(int Row, int Column, int RowSpan, int ColumnSpan);
