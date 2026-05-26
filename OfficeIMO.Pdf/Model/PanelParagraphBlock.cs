namespace OfficeIMO.Pdf;

internal sealed class PanelParagraphBlock : IPdfBlock {
    public System.Collections.Generic.IReadOnlyList<TextRun> Runs { get; }
    public PdfAlign Align { get; }
    public PdfColor? DefaultColor { get; }
    public PanelStyle? Style { get; }
    public PanelParagraphBlock(System.Collections.Generic.IEnumerable<TextRun> runs, PdfAlign align, PdfColor? defaultColor, PanelStyle? style = null) {
        Guard.NotNull(runs, nameof(runs));
        Guard.ParagraphAlign(align, nameof(align), "Panel paragraph");
        var snapshot = new System.Collections.Generic.List<TextRun>();
        snapshot.AddRange(runs);
        Align = align; DefaultColor = defaultColor; Style = style?.Clone(); Runs = snapshot.AsReadOnly();
    }
}
