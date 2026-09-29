namespace OfficeIMO.Pdf;

/// <summary>Semantic or visual change detected between an aligned page pair.</summary>
public enum PdfReviewChangeKind {
    /// <summary>Readable text appears only on the actual page.</summary>
    TextAdded,
    /// <summary>Readable text appears only on the expected page.</summary>
    TextRemoved,
    /// <summary>Readable text changed at a corresponding position.</summary>
    TextChanged,
    /// <summary>Identical readable text changed position.</summary>
    TextMoved,
    /// <summary>An image placement appears only on the actual page.</summary>
    ImageAdded,
    /// <summary>An image placement appears only on the expected page.</summary>
    ImageRemoved,
    /// <summary>Image payloads differ at corresponding placements; visual review remains necessary.</summary>
    ImageChangedCandidate,
    /// <summary>The same image payload changed placement position or displayed size.</summary>
    ImageMoved,
    /// <summary>A scanned or image-dominated page differs and requires visual or OCR review.</summary>
    ScannedPageUncertain,
    /// <summary>Rendered pixels differ without a supported semantic classification.</summary>
    UnclassifiedVisual,
    /// <summary>Managed rendering skipped or approximated page content, so equality cannot be proved.</summary>
    RenderUncertain
}

/// <summary>One page-linked change with optional top-left visual geometry and text evidence.</summary>
public sealed class PdfReviewChange {
    internal PdfReviewChange(PdfReviewChangeKind kind, int expectedPageNumber, int actualPageNumber, PdfLogicalVisualBounds? expectedBounds, PdfLogicalVisualBounds? actualBounds, string? expectedText = null, string? actualText = null, bool canCoverRenderedPixels = true) {
        Kind = kind;
        ExpectedPageNumber = expectedPageNumber;
        ActualPageNumber = actualPageNumber;
        ExpectedBounds = expectedBounds;
        ActualBounds = actualBounds;
        ExpectedText = expectedText;
        ActualText = actualText;
        CanCoverRenderedPixels = canCoverRenderedPixels;
    }
    /// <summary>Classification, including explicit uncertainty for scans and unsupported visual changes.</summary>
    public PdfReviewChangeKind Kind { get; }
    /// <summary>One-based expected page.</summary>
    public int ExpectedPageNumber { get; }
    /// <summary>One-based actual page.</summary>
    public int ActualPageNumber { get; }
    /// <summary>Expected-page geometry in top-left visual points, when available.</summary>
    public PdfLogicalVisualBounds? ExpectedBounds { get; }
    /// <summary>Actual-page geometry in top-left visual points, when available.</summary>
    public PdfLogicalVisualBounds? ActualBounds { get; }
    /// <summary>Previous readable text, when relevant.</summary>
    public string? ExpectedText { get; }
    /// <summary>Current readable text, when relevant.</summary>
    public string? ActualText { get; }
    internal bool CanCoverRenderedPixels { get; }
}

/// <summary>One aligned pair with retained rendered proof and navigable content changes.</summary>
public sealed class PdfReviewPageComparison {
    internal PdfReviewPageComparison(PdfPageChange alignment, PdfVisualPageComparison? visual, IReadOnlyList<PdfReviewChange> changes) {
        ExpectedPageNumber = alignment.ExpectedPageNumber!.Value;
        ActualPageNumber = alignment.ActualPageNumber!.Value;
        AlignmentKind = alignment.Kind;
        Visual = visual;
        Changes = Array.AsReadOnly(changes.ToArray());
    }
    /// <summary>One-based expected source page.</summary>
    public int ExpectedPageNumber { get; }
    /// <summary>One-based actual source page.</summary>
    public int ActualPageNumber { get; }
    /// <summary>Rendered alignment relationship for this page pair.</summary>
    public PdfPageChangeKind AlignmentKind { get; }
    /// <summary>Rendered expected, actual, and highlighted difference images for a modified candidate; null when exact rendered alignment already proved the images equal.</summary>
    public PdfVisualPageComparison? Visual { get; }
    /// <summary>Classified changes in page reading order.</summary>
    public IReadOnlyList<PdfReviewChange> Changes { get; }
    /// <summary>True only when both rendered and supported semantic content match.</summary>
    public bool IsMatch => (Visual?.IsMatch ?? true) && Changes.Count == 0;
}

/// <summary>Bounded comparison settings that reuse the existing alignment and visual contracts.</summary>
public sealed class PdfReviewComparisonOptions {
    /// <summary>Rendered-page matching settings. RenderScale and Background must match Visual settings.</summary>
    public PdfPageChangeOptions PageAlignment { get; set; } = new PdfPageChangeOptions();
    /// <summary>Detailed visual settings. Scale and Background must match PageAlignment settings. A pixel ignore region suppresses semantic text or image evidence only when it fully contains that element's bounds.</summary>
    public PdfVisualComparisonOptions Visual { get; set; } = new PdfVisualComparisonOptions();
    /// <summary>Maximum changed page pairs analyzed semantically and visually.</summary>
    public int MaxChangedPagePairs { get; set; } = 50;
    /// <summary>Maximum aligned page pairs inspected for semantic differences, including exact rendered matches.</summary>
    public int MaxAlignedPagePairs { get; set; } = 100;
    /// <summary>Maximum text blocks inspected on either page of a pair.</summary>
    public int MaxTextBlocksPerPage { get; set; } = 300;
    /// <summary>Maximum image placements inspected on either page of a pair.</summary>
    public int MaxImagePlacementsPerPage { get; set; } = 100;

    internal void Validate() {
        Guard.NotNull(PageAlignment, nameof(PageAlignment));
        Guard.NotNull(Visual, nameof(Visual));
        PageAlignment.Validate();
        Visual.Validate();
        // An exact alignment fingerprint can replace detailed rendering only under the same raster policy.
        if (PageAlignment.RenderScale != Visual.Scale || !PageAlignment.Background.Equals(Visual.Background)) {
            throw new ArgumentException("Page alignment and visual comparison must use the same render scale and background.");
        }
        if (MaxChangedPagePairs <= 0) throw new ArgumentOutOfRangeException(nameof(MaxChangedPagePairs));
        if (MaxAlignedPagePairs <= 0) throw new ArgumentOutOfRangeException(nameof(MaxAlignedPagePairs));
        if (MaxTextBlocksPerPage <= 0) throw new ArgumentOutOfRangeException(nameof(MaxTextBlocksPerPage));
        if (MaxImagePlacementsPerPage <= 0) throw new ArgumentOutOfRangeException(nameof(MaxImagePlacementsPerPage));
    }
}

/// <summary>Page alignment and detailed review results for changed page pairs.</summary>
public sealed class PdfReviewComparisonReport {
    internal PdfReviewComparisonReport(PdfPageChangeReport alignment, IReadOnlyList<PdfReviewPageComparison> pages) {
        PageAlignment = alignment;
        Pages = Array.AsReadOnly(pages.ToArray());
    }
    /// <summary>Inserted, deleted, moved, unchanged, and modified-candidate page relationships.</summary>
    public PdfPageChangeReport PageAlignment { get; }
    /// <summary>Paired pages with visual or semantic differences; exact rendered pairs can contain semantic-only changes.</summary>
    public IReadOnlyList<PdfReviewPageComparison> Pages { get; }
    /// <summary>True when page order and every reviewed page pair match under the selected visual policy.</summary>
    public bool IsMatch => PageAlignment.Changes.All(static change => change.Kind is PdfPageChangeKind.Unchanged or PdfPageChangeKind.ModifiedCandidate) &&
        Pages.All(static page => page.IsMatch);
}
