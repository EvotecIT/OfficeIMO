namespace OfficeIMO.Pdf;

internal sealed class HeadingBlock : IPdfBlock {
    public int Level { get; }
    public string Text { get; }
    public PdfAlign Align { get; }
    public PdfColor? Color { get; }
    public string? LinkUri { get; }
    public string? LinkDestinationName { get; }
    public string? LinkContents { get; }
    public PdfHeadingStyle? Style { get; }
    public HeadingBlock(int level, string text, PdfAlign align, PdfColor? color, string? linkUri = null, PdfHeadingStyle? style = null, string? linkContents = null, string? linkDestinationName = null) {
        Guard.NotNullOrWhiteSpace(text, nameof(text));
        Guard.LeftCenterRightAlign(align, nameof(align), "Heading");
        if (linkUri != null && linkDestinationName != null) {
            throw new System.ArgumentException("A heading link can target either a URI or a bookmark, not both.", nameof(linkDestinationName));
        }

        bool hasLinkTarget = linkUri != null || linkDestinationName != null;
        if (linkContents != null && !hasLinkTarget) {
            throw new System.ArgumentException("Heading link annotation contents require a link target.", nameof(linkContents));
        }

        if (linkUri != null) {
            Guard.UriAction(linkUri, nameof(linkUri));
        }

        if (linkDestinationName != null) {
            Guard.NotNullOrWhiteSpace(linkDestinationName, nameof(linkDestinationName));
        }

        if (hasLinkTarget && linkContents != null) {
            Guard.NotNullOrWhiteSpace(linkContents, nameof(linkContents));
        }

        Level = level; Text = text; Align = align; Color = color; LinkUri = linkUri; LinkDestinationName = linkDestinationName; LinkContents = hasLinkTarget ? linkContents ?? text : null; Style = style?.Clone();
    }
}
