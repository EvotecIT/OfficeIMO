namespace OfficeIMO.Pdf;

internal sealed class RichParagraphBlock : IPdfBlock {
    public System.Collections.Generic.IReadOnlyList<PdfTextRun> Runs { get; }
    public PdfAlign Align { get; }
    public global::OfficeIMO.OfficeColor? DefaultColor { get; }
    public PdfParagraphStyle? Style { get; }
    public RichParagraphBlock(System.Collections.Generic.IEnumerable<PdfTextRun> runs, PdfAlign align, global::OfficeIMO.OfficeColor? defaultColor, PdfParagraphStyle? style = null) {
        Guard.NotNull(runs, nameof(runs));
        Guard.ParagraphAlign(align, nameof(align), "Paragraph");
        var snapshot = new System.Collections.Generic.List<PdfTextRun>();
        snapshot.AddRange(runs);
        Align = align; DefaultColor = defaultColor; Style = style?.Clone(); Runs = snapshot.AsReadOnly();
    }
}
