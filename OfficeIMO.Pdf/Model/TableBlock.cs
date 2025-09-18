namespace OfficeIMO.Pdf;

internal sealed class TableBlock : IPdfBlock {
    public System.Collections.Generic.List<string[]> Rows { get; } = new();
    public PdfAlign Align { get; }
    public PdfTableStyle? Style { get; }
    public TableBlock(System.Collections.Generic.IEnumerable<string[]> rows, PdfAlign align, PdfTableStyle? style) { Align = align; Style = style; Rows.AddRange(rows); }
}

public sealed class PdfTableStyle {
    public PdfColor? BorderColor { get; set; } = new PdfColor(0.8, 0.8, 0.8);
    public double BorderWidth { get; set; } = 0.5;
    public PdfColor? HeaderFill { get; set; } = new PdfColor(0.95, 0.95, 0.95);
    public PdfColor? RowStripeFill { get; set; } = new PdfColor(0.98, 0.98, 0.98);
    public PdfColor? TextColor { get; set; }
    public PdfColor? HeaderTextColor { get; set; }
    public double CellPaddingX { get; set; } = 4;
    public double CellPaddingY { get; set; } = 2;
    /// <summary>Vertical baseline adjustment (points). Positive moves text down, negative up.</summary>
    public double RowBaselineOffset { get; set; } = 0;
    /// <summary>Optional per-column alignment (falls back to Left).</summary>
    public System.Collections.Generic.List<PdfColumnAlign>? Alignments { get; set; }
    /// <summary>When true, cells that look numeric are right-aligned automatically.</summary>
    public bool RightAlignNumeric { get; set; }
}
