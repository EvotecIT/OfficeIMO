namespace OfficeIMO.Pdf;

internal sealed class RichParagraphBlock : IPdfBlock {
    public System.Collections.Generic.IReadOnlyList<PdfTextRun> Runs { get; }
    public OfficeAlignment Align { get; }
    public global::OfficeIMO.OfficeColor? DefaultColor { get; }
    public PdfParagraphStyle? Style { get; }
    /// <summary>The live authoring paragraph whose runs this block renders, when it came from one.</summary>
    internal PdfParagraph? Source { get; }

    public RichParagraphBlock(System.Collections.Generic.IEnumerable<PdfTextRun> runs, OfficeAlignment align, global::OfficeIMO.OfficeColor? defaultColor, PdfParagraphStyle? style = null) {
        Guard.NotNull(runs, nameof(runs));
        align = Guard.ParagraphAlign(align, nameof(align), "Paragraph");
        var snapshot = new System.Collections.Generic.List<PdfTextRun>();
        snapshot.AddRange(runs);
        Align = align; DefaultColor = defaultColor; Style = style?.Clone(); Runs = snapshot.AsReadOnly();
    }

    // Shares the paragraph's run list so later AddText/Set* calls on the returned paragraph are rendered.
    internal RichParagraphBlock(PdfParagraph paragraph, PdfParagraphStyle? style) {
        Source = paragraph;
        Align = paragraph.Alignment; DefaultColor = paragraph.FallbackColor; Style = style?.Clone(); Runs = paragraph.Runs;
    }

    private RichParagraphBlock(RichParagraphBlock source, PdfParagraphStyle? style) {
        Source = source.Source;
        Align = source.Align; DefaultColor = source.DefaultColor; Style = style; Runs = source.Runs;
    }

    /// <summary>Same runs (still live when authored through a paragraph) with another layout style.</summary>
    internal RichParagraphBlock WithStyle(PdfParagraphStyle? style) => new RichParagraphBlock(this, style?.Clone());
}
