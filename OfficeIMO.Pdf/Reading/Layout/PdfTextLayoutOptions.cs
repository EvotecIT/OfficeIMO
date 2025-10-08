namespace OfficeIMO.Pdf;

/// <summary>
/// Public options for column-aware text extraction.
/// </summary>
public sealed class PdfTextLayoutOptions {
    /// <summary>Left page margin in points used when inferring column bounds. Default: 36.</summary>
    public double MarginLeft { get; set; } = 36;
    /// <summary>Right page margin in points used when inferring column bounds. Default: 36.</summary>
    public double MarginRight { get; set; } = 36;
    /// <summary>Histogram bin width (points) used for gutter detection. Default: 5.</summary>
    public double BinWidth { get; set; } = 5;
    /// <summary>Minimum detected gutter width (points) to consider a two-column layout. Default: 24.</summary>
    public double MinGutterWidth { get; set; } = 24;
    /// <summary>Maximum Y delta, expressed in font-size em units, to merge spans into the same text line. Default: 0.6.</summary>
    public double LineMergeToleranceEm { get; set; } = 0.6;
    /// <summary>Maximum absolute Y delta (points) to merge spans into the same line. Default: 2.5.</summary>
    public double LineMergeMaxPoints { get; set; } = 2.5;
    /// <summary>When true, forces single-column reading order and disables gutter detection. Default: false.</summary>
    public bool ForceSingleColumn { get; set; }
    /// <summary>When true, joins hyphenated words broken across line ends. Default: true.</summary>
    public bool JoinHyphenationAcrossLines { get; set; } = true;
    /// <summary>Height from top of page (points) to ignore as header when emitting text. Default: 0.</summary>
    public double IgnoreHeaderHeight { get; set; }
    /// <summary>Height from bottom of page (points) to ignore as footer when emitting text. Default: 0.</summary>
    public double IgnoreFooterHeight { get; set; }
    /// <summary>Threshold in em units to insert a space between adjacent spans on the same line. Default: 0.3.</summary>
    public double GapSpaceThresholdEm { get; set; } = 0.3;
    /// <summary>Threshold as a fraction of previous span's average glyph advance to insert a space. Default: 0.45.</summary>
    public double GapGlyphFactor { get; set; } = 0.45;

    internal TextLayoutEngine.Options ToEngineOptions() => new TextLayoutEngine.Options {
        MarginLeft = this.MarginLeft,
        MarginRight = this.MarginRight,
        BinWidth = this.BinWidth,
        MinGutterWidth = this.MinGutterWidth,
        LineMergeToleranceEm = this.LineMergeToleranceEm,
        LineMergeMaxPoints = this.LineMergeMaxPoints,
        ForceSingleColumn = this.ForceSingleColumn,
        GapSpaceThresholdEm = this.GapSpaceThresholdEm,
        GapGlyphFactor = this.GapGlyphFactor
    };
}
