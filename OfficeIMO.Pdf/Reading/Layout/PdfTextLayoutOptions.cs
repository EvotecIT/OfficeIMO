namespace OfficeIMO.Pdf;

/// <summary>Preferred horizontal reading direction for ambiguous positioned text.</summary>
public enum PdfReadingDirection {
    /// <summary>Use the first strong Unicode directional character, falling back to geometric left-to-right order.</summary>
    Auto,
    /// <summary>Read horizontal baselines and columns from left to right.</summary>
    LeftToRight,
    /// <summary>Read horizontal baselines and columns from right to left.</summary>
    RightToLeft
}

/// <summary>
/// Public options for column-aware text extraction.
/// </summary>
public sealed class PdfTextLayoutOptions {
    /// <summary>Left page margin in points used when inferring column bounds. Default: 36.</summary>
    public OfficeLength MarginLeft { get; set; } = OfficeIMO.OfficeLength.Points(36);
    /// <summary>Right page margin in points used when inferring column bounds. Default: 36.</summary>
    public OfficeLength MarginRight { get; set; } = OfficeIMO.OfficeLength.Points(36);
    /// <summary>Histogram bin width (points) used for gutter detection. Default: 5.</summary>
    public OfficeLength BinWidth { get; set; } = OfficeIMO.OfficeLength.Points(5);
    /// <summary>Minimum detected gutter width (points) to consider a two-column layout. Default: 24.</summary>
    public OfficeLength MinGutterWidth { get; set; } = OfficeIMO.OfficeLength.Points(24);
    /// <summary>Maximum baseline-normal delta in font-size em units for ordinary line clustering. Default: 0.6. A stricter value disables conservative inline-script attachment.</summary>
    public double LineMergeToleranceEm { get; set; } = 0.6;
    /// <summary>Maximum absolute baseline-normal delta for ordinary line clustering. Default: 2.5 points. With both merge limits at least their defaults, adjacent smaller-font scripts may attach to a unique base line; a stricter value retains baseline-only grouping.</summary>
    public OfficeLength LineMergeMax { get; set; } = OfficeIMO.OfficeLength.Points(2.5);
    /// <summary>When true, forces single-column reading order and disables gutter detection. Default: false.</summary>
    public bool ForceSingleColumn { get; set; }
    /// <summary>Horizontal reading direction used for words and column partitions. Default: Auto.</summary>
    public PdfReadingDirection ReadingDirection { get; set; }
    /// <summary>
    /// When true, removes an explicit soft hyphen before a line break. Visible hyphens are
    /// preserved because PDF layout and letter case cannot prove that they are discretionary.
    /// Default: false.
    /// </summary>
    public bool JoinSoftHyphensAcrossLines { get; set; }
    /// <summary>Height from top of page (points) to ignore as header when emitting text. Default: 0.</summary>
    public OfficeLength IgnoreHeaderHeight { get; set; }
    /// <summary>Height from bottom of page (points) to ignore as footer when emitting text. Default: 0.</summary>
    public OfficeLength IgnoreFooterHeight { get; set; }
    /// <summary>Threshold in em units to insert a space between adjacent spans on the same line. Default: 0.35.</summary>
    public double GapSpaceThresholdEm { get; set; } = 0.35;
    /// <summary>Threshold as a fraction of previous span's average glyph advance to insert a space. Default: 0.60.</summary>
    public double GapGlyphFactor { get; set; } = 0.60;

    internal TextLayoutEngine.Options ToEngineOptions() => new TextLayoutEngine.Options {
        MarginLeft = this.MarginLeft.ToPoints(),
        MarginRight = this.MarginRight.ToPoints(),
        BinWidth = this.BinWidth.ToPoints(),
        MinGutterWidth = this.MinGutterWidth.ToPoints(),
        LineMergeToleranceEm = this.LineMergeToleranceEm,
        LineMergeMaxPoints = this.LineMergeMax.ToPoints(),
        ForceSingleColumn = this.ForceSingleColumn,
        ReadingDirection = this.ReadingDirection,
        GapSpaceThresholdEm = this.GapSpaceThresholdEm,
        GapGlyphFactor = this.GapGlyphFactor
    };
}
