namespace OfficeIMO.Pdf;

public sealed partial class PdfTextRun {
    /// <summary>The page-associated footnote referenced by this run, or null.</summary>
    public PdfFootnoteDefinition? FootnoteReference { get; private set; }

    /// <summary>Creates a copy associated with a footnote while retaining text, typography, spacing, and links.</summary>
    /// <remarks>The text is the author's visible reference label and may be empty. The associated note begins on the page admitting this run.</remarks>
    public PdfTextRun WithFootnoteReference(PdfFootnoteDefinition definition) {
        Guard.NotNull(definition, nameof(definition));
        if (InlineElement != null)
            throw new ArgumentException("A footnote reference must be an ordinary text run without layout controls.", nameof(definition));
        foreach (char character in Text) {
            if (character is '\r' or '\n' or '\t')
                throw new ArgumentException("A footnote reference must be an ordinary text run without layout controls.", nameof(definition));
        }
        var copy = CopyRunFormattingTo(new PdfTextRun(Text, Bold, Underline, Color, Italic, Strike,
            FontSize, Font, LinkUri, LinkContents, Baseline, LinkDestinationName, TabLeader, TabAlignment,
            BackgroundColor, FontFamily, UnderlineStyle, StrikeStyle, DecorationColor));
        copy.FootnoteReference = definition;
        return copy;
    }

}
