using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

/// <summary>
/// Positioned image retained in its encoded source form for image and PDF backends.
/// </summary>
public sealed class HtmlRenderImage : HtmlRenderVisual {
    private readonly HtmlRenderImageData _data;

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
        string? source = null,
        OfficeImageSourceCrop sourceCrop = default)
        : this(new HtmlRenderImageData(bytes), contentType, x, y, width, height, paintOrder, alternativeText, linkUri, source, sourceCrop) {
    }

    private HtmlRenderImage(
        HtmlRenderImageData data,
        string contentType,
        double x,
        double y,
        double width,
        double height,
        int paintOrder,
        string? alternativeText,
        string? linkUri,
        string? source,
        OfficeImageSourceCrop sourceCrop)
        : base(HtmlRenderVisualKind.Image, x, y, width, height, paintOrder, linkUri, source) {
        _data = data ?? throw new ArgumentNullException(nameof(data));
        ContentType = string.IsNullOrWhiteSpace(contentType) ? "application/octet-stream" : contentType;
        AlternativeText = alternativeText;
        SourceCrop = sourceCrop;
    }

    /// <summary>Detached encoded image bytes.</summary>
    public byte[] Bytes => _data.Snapshot();

    /// <summary>Image media type.</summary>
    public string ContentType { get; }

    /// <summary>Optional alternative text from the source document.</summary>
    public string? AlternativeText { get; }

    /// <summary>Optional normalized source crop applied before the image is placed.</summary>
    public OfficeImageSourceCrop SourceCrop { get; }

    internal byte[] EncodedBytes => _data.EncodedBytes;

    internal override HtmlRenderVisual Translate(double offsetX, double offsetY, int paintOrder) =>
        new HtmlRenderImage(_data, ContentType, X + offsetX, Y + offsetY, Width, Height, paintOrder, AlternativeText, LinkUri, Source, SourceCrop);
}
