using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal sealed class ShapeBlock : IPdfBlock {
    public OfficeShape Shape { get; }
    public PdfDrawingStyle? Style { get; }
    public string? LinkUri { get; }
    public string? LinkContents { get; }
    public PdfAlign Align => (Style ?? new PdfDrawingStyle()).Align;
    public double SpacingBefore => (Style ?? new PdfDrawingStyle()).SpacingBefore;
    public double SpacingAfter => (Style ?? new PdfDrawingStyle()).SpacingAfter;
    public string? AlternativeText => Style?.AlternativeText;
    public bool Decorative => Style?.Decorative == true;

    public ShapeBlock(OfficeShape shape, PdfDrawingStyle? style = null, string? linkUri = null, string? linkContents = null) {
        Guard.NotNull(shape, nameof(shape));
        Guard.OptionalUriAction(linkUri, nameof(linkUri));
        if (linkContents != null && linkUri == null) {
            throw new ArgumentException("Shape link contents require a link URI.", nameof(linkContents));
        }

        if (linkContents != null) {
            Guard.NotNullOrWhiteSpace(linkContents, nameof(linkContents));
        }

        Shape = shape.Clone();
        Style = style?.Clone();
        LinkUri = linkUri;
        LinkContents = linkUri == null ? null : linkContents ?? "Shape";
    }

    public ShapeBlock(OfficeShape shape, PdfAlign align, double spacingBefore, double spacingAfter, string? linkUri = null, string? linkContents = null)
        : this(shape, new PdfDrawingStyle {
            Align = align,
            SpacingBefore = spacingBefore,
            SpacingAfter = spacingAfter
        }, linkUri, linkContents) {
    }
}
