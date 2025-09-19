namespace OfficeIMO.Pdf;

public sealed class PanelStyle {
    public PdfColor? Background { get; set; }
    public PdfColor? BorderColor { get; set; }
    public double BorderWidth { get; set; } = 0.5;
    public double PaddingY { get; set; } = 6; // vertical padding
    public double PaddingX { get; set; } = 6; // horizontal padding
    public double? MaxWidth { get; set; } // when set, panel box is centered/narrowed
    public PdfAlign Align { get; set; } = PdfAlign.Left; // alignment of the panel box within content width
}

internal sealed class PanelParagraphBlock : IPdfBlock {
    public System.Collections.Generic.List<TextRun> Runs { get; } = new();
    public PdfAlign Align { get; }
    public PdfColor? DefaultColor { get; }
    public PanelStyle Style { get; }
    public PanelParagraphBlock(System.Collections.Generic.IEnumerable<TextRun> runs, PdfAlign align, PdfColor? defaultColor, PanelStyle style) {
        Align = align; DefaultColor = defaultColor; Style = style; Runs.AddRange(runs);
    }
}
