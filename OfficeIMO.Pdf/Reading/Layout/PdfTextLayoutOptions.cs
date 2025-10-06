namespace OfficeIMO.Pdf;

/// <summary>
/// Public options for column-aware text extraction.
/// </summary>
public sealed class PdfTextLayoutOptions {
    public double MarginLeft { get; set; } = 36;
    public double MarginRight { get; set; } = 36;
    public double BinWidth { get; set; } = 5;
    public double MinGutterWidth { get; set; } = 24;
    public double LineMergeToleranceEm { get; set; } = 0.6;
    public bool ForceSingleColumn { get; set; } = false;
    public bool JoinHyphenationAcrossLines { get; set; } = true;

    internal TextLayoutEngine.Options ToEngineOptions() => new TextLayoutEngine.Options {
        MarginLeft = this.MarginLeft,
        MarginRight = this.MarginRight,
        BinWidth = this.BinWidth,
        MinGutterWidth = this.MinGutterWidth,
        LineMergeToleranceEm = this.LineMergeToleranceEm,
        ForceSingleColumn = this.ForceSingleColumn
    };
}
