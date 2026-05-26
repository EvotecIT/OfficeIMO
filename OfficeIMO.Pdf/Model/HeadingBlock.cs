namespace OfficeIMO.Pdf;

internal sealed class HeadingBlock : IPdfBlock {
    public int Level { get; }
    public string Text { get; }
    public PdfAlign Align { get; }
    public PdfColor? Color { get; }
    public string? LinkUri { get; }
    public string? LinkContents { get; }
    public PdfHeadingStyle? Style { get; }
    public HeadingBlock(int level, string text, PdfAlign align, PdfColor? color, string? linkUri = null, PdfHeadingStyle? style = null, string? linkContents = null) {
        Guard.NotNullOrWhiteSpace(text, nameof(text));
        Guard.LeftCenterRightAlign(align, nameof(align), "Heading");
        if (linkContents != null && linkUri == null) {
            throw new System.ArgumentException("Heading link annotation contents require a heading link URI.", nameof(linkContents));
        }

        if (linkUri != null) {
            Guard.AbsoluteUri(linkUri, nameof(linkUri));
            if (linkContents != null) {
                Guard.NotNullOrWhiteSpace(linkContents, nameof(linkContents));
            }
        }

        Level = level; Text = text; Align = align; Color = color; LinkUri = linkUri; LinkContents = linkUri == null ? null : linkContents ?? text; Style = style?.Clone();
    }
}
