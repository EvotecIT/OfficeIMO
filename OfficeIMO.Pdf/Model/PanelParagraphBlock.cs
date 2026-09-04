namespace OfficeIMO.Pdf;

internal sealed class PanelParagraphBlock : IPdfBlock {
    public System.Collections.Generic.IReadOnlyList<PdfTextRun> Runs { get; }
    public PdfAlign Align { get; }
    public PdfColor? DefaultColor { get; }
    public PdfPanelStyle? Style { get; }
    public PanelParagraphBlock(System.Collections.Generic.IEnumerable<PdfTextRun> runs, PdfAlign align, PdfColor? defaultColor, PdfPanelStyle? style = null) {
        Guard.NotNull(runs, nameof(runs));
        Guard.ParagraphAlign(align, nameof(align), "Panel paragraph");
        var snapshot = new System.Collections.Generic.List<PdfTextRun>();
        snapshot.AddRange(runs);
        Align = align; DefaultColor = defaultColor; Style = style?.Clone(); Runs = snapshot.AsReadOnly();
    }
}
