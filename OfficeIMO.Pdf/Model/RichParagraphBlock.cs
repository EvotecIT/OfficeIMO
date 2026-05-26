namespace OfficeIMO.Pdf;

internal sealed class RichParagraphBlock : IPdfBlock {
    public System.Collections.Generic.IReadOnlyList<TextRun> Runs { get; }
    public PdfAlign Align { get; }
    public PdfColor? DefaultColor { get; }
    public PdfParagraphStyle? Style { get; }
    public RichParagraphBlock(System.Collections.Generic.IEnumerable<TextRun> runs, PdfAlign align, PdfColor? defaultColor, PdfParagraphStyle? style = null) {
        Guard.NotNull(runs, nameof(runs));
        Guard.ParagraphAlign(align, nameof(align), "Paragraph");
        var snapshot = new System.Collections.Generic.List<TextRun>();
        snapshot.AddRange(runs);
        Align = align; DefaultColor = defaultColor; Style = style?.Clone(); Runs = snapshot.AsReadOnly();
    }
}
