namespace OfficeIMO.Pdf;

internal sealed class ParagraphBlock : IPdfBlock {
    public string Text { get; }
    public ParagraphBlock(string text) { Text = text; }
}

