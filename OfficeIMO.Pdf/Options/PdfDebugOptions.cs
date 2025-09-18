namespace OfficeIMO.Pdf;

/// <summary>
/// Debug drawing toggles. When enabled, the writer overlays guides (margins, baselines, row boxes).
/// </summary>
public sealed class PdfDebugOptions {
    /// <summary>Draw the content area (margins) rectangle on each page.</summary>
    public bool ShowContentArea { get; set; }
    /// <summary>Draw a thin baseline for each table row.</summary>
    public bool ShowTableBaselines { get; set; }
    /// <summary>Draw row rectangles (independent of styling) to verify box alignment.</summary>
    public bool ShowTableRowBoxes { get; set; }
    /// <summary>Draw table column guides (vertical lines at column boundaries).</summary>
    public bool ShowTableColumnGuides { get; set; }
}

