namespace OfficeIMO.Pdf;

/// <summary>How a page relates to a second document at the selected render scale.</summary>
public enum PdfPageChangeKind {
    /// <summary>The rendered pixels match, and relative page order is preserved.</summary>
    Unchanged,
    /// <summary>The rendered pixels match, but relative page order changed.</summary>
    Moved,
    /// <summary>Pages occupy the same unmatched sequence slot and need visual review.</summary>
    ModifiedCandidate,
    /// <summary>A page appears only in the actual document.</summary>
    Inserted,
    /// <summary>A page appears only in the expected document.</summary>
    Deleted
}

/// <summary>One page alignment. Null page numbers identify insertion or deletion.</summary>
public sealed class PdfPageChange {
    internal PdfPageChange(PdfPageChangeKind kind, int? expectedPageNumber, int? actualPageNumber, bool usesIgnoredRegions = false) {
        Kind = kind;
        ExpectedPageNumber = expectedPageNumber;
        ActualPageNumber = actualPageNumber;
        UsesIgnoredRegions = usesIgnoredRegions;
    }

    /// <summary>Relationship between the pages.</summary>
    public PdfPageChangeKind Kind { get; }
    /// <summary>One-based page in the expected document, when present.</summary>
    public int? ExpectedPageNumber { get; }
    /// <summary>One-based page in the actual document, when present.</summary>
    public int? ActualPageNumber { get; }
    /// <summary>Whether ignored pixel regions contributed to this alignment.</summary>
    public bool UsesIgnoredRegions { get; }
    /// <summary>True only when all rendered pixels match at the configured scale without masked regions.</summary>
    public bool IsExactRenderedMatch => !UsesIgnoredRegions && (Kind == PdfPageChangeKind.Unchanged || Kind == PdfPageChangeKind.Moved);
}

/// <summary>Bounded page alignment for review. Changes remain candidates until inspected at full fidelity.</summary>
public sealed class PdfPageChangeReport {
    internal PdfPageChangeReport(IReadOnlyList<PdfPageChange> changes, int expectedPageCount, int actualPageCount, double renderScale) {
        Changes = Array.AsReadOnly(changes.ToArray());
        ExpectedPageCount = expectedPageCount;
        ActualPageCount = actualPageCount;
        RenderScale = renderScale;
    }

    /// <summary>Expected document page count.</summary>
    public int ExpectedPageCount { get; }
    /// <summary>Actual document page count.</summary>
    public int ActualPageCount { get; }
    /// <summary>Scale used for page raster matching.</summary>
    public double RenderScale { get; }
    /// <summary>Expected-page order followed by actual-only inserted pages.</summary>
    public IReadOnlyList<PdfPageChange> Changes { get; }
    /// <summary>Whether every page matched and retained its relative order.</summary>
    public bool IsMatch => ExpectedPageCount == ActualPageCount && Changes.All(static change => change.Kind == PdfPageChangeKind.Unchanged);
}
