namespace OfficeIMO.Html;

/// <summary>
/// Positioned image retained in its encoded source form for image and PDF backends.
/// </summary>
public sealed class HtmlRenderImage : HtmlRenderVisual {
    internal HtmlRenderImage(
        byte[] bytes,
        string contentType,
        double x,
        double y,
        double width,
        double height,
        int paintOrder,
        string? alternativeText = null,
        string? linkUri = null,
        string? source = null)
        : base(HtmlRenderVisualKind.Image, x, y, width, height, paintOrder, linkUri, source) {
        if (bytes == null || bytes.Length == 0) {
            throw new ArgumentException("Rendered images require encoded bytes.", nameof(bytes));
        }

        Bytes = (byte[])bytes.Clone();
        ContentType = string.IsNullOrWhiteSpace(contentType) ? "application/octet-stream" : contentType;
        AlternativeText = alternativeText;
    }

    /// <summary>Encoded image bytes.</summary>
    public byte[] Bytes { get; }

    /// <summary>Image media type.</summary>
    public string ContentType { get; }

    /// <summary>Optional alternative text from the source document.</summary>
    public string? AlternativeText { get; }

    internal override HtmlRenderVisual Translate(double offsetX, double offsetY, int paintOrder) =>
        new HtmlRenderImage(Bytes, ContentType, X + offsetX, Y + offsetY, Width, Height, paintOrder, AlternativeText, LinkUri, Source);
}
