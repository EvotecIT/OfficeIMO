namespace OfficeIMO.Pdf;

public sealed partial class PdfTextRun {
    /// <summary>Retains run metadata when the layout projects an inline visual into another text frame.</summary>
    internal PdfTextRun WithInlineElement(PdfInlineElement element) => new PdfTextRun(this, element);

    private PdfTextRun(PdfTextRun source, PdfInlineElement element)
        : this(source.Text, source.Bold, source.Underline, source.Color, source.Italic, source.Strike,
            source.FontSize, source.Font, source.LinkUri, source.LinkContents, source.Baseline,
            source.LinkDestinationName, source.TabLeader, source.TabAlignment, source.BackgroundColor,
            source.FontFamily, source.UnderlineStyle, source.StrikeStyle, source.DecorationColor) {
        InlineElement = element;
        source.CopyRunFormattingTo(this);
    }
}
