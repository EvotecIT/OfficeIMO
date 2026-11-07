using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal sealed class DrawingBlock : IPdfBlock {
    public OfficeDrawing Drawing { get; }
    public PdfDrawingStyle? Style { get; }
    public string? LinkUri { get; }
    public string? LinkContents { get; }
    public OfficeAlignment Align => (Style ?? new PdfDrawingStyle()).Align;
    public double SpacingBefore => (Style ?? new PdfDrawingStyle()).SpacingBefore.ToPoints();
    public double SpacingAfter => (Style ?? new PdfDrawingStyle()).SpacingAfter.ToPoints();
    public string? AlternativeText => Style?.AlternativeText;
    public bool Decorative => Style?.Decorative == true;

    /// <summary>Retains the factory's private validated style. Shared scene snapshots must remain immutable and unavailable to caller edits.</summary>
    internal DrawingBlock(OfficeDrawing drawing, PdfDrawingStyle? preparedStyle, string? linkUri, string? linkContents, bool useDrawingSnapshot) {
        Guard.NotNull(drawing, nameof(drawing));
        Guard.OptionalUriAction(linkUri, nameof(linkUri));
        if (linkContents != null && linkUri == null) {
            throw new ArgumentException("Drawing link contents require a link URI.", nameof(linkContents));
        }

        if (linkContents != null) {
            Guard.NotNullOrWhiteSpace(linkContents, nameof(linkContents));
        }

        Drawing = useDrawingSnapshot ? drawing : drawing.Clone();
        Style = preparedStyle;
        LinkUri = linkUri;
        LinkContents = linkUri == null ? null : linkContents ?? "Drawing";
    }
}
