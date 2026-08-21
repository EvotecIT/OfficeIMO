using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

/// <summary>Positioned shared vector drawing emitted by HTML layout.</summary>
public sealed class HtmlRenderDrawing : HtmlRenderVisual {
    private readonly OfficeDrawing _drawing;
    private readonly HtmlRenderImageData? _imageData;
    private readonly double? _imageX;
    private readonly double? _imageY;
    private readonly double? _imageWidth;
    private readonly double? _imageHeight;

    internal HtmlRenderDrawing(
        OfficeDrawing drawing,
        double x,
        double y,
        double width,
        double height,
        int paintOrder,
        string? alternativeText,
        string? linkUri,
        string? source,
        double? layoutY = null,
        byte[]? imageBytes = null,
        string? imageContentType = null,
        OfficeImageSourceCrop sourceCrop = default,
        double? imageX = null,
        double? imageY = null,
        double? imageWidth = null,
        double? imageHeight = null)
        : base(HtmlRenderVisualKind.Drawing, x, y, width, height, paintOrder, linkUri, source, layoutY) {
        _drawing = (drawing ?? throw new ArgumentNullException(nameof(drawing))).Clone();
        _imageData = imageBytes is { Length: > 0 } ? new HtmlRenderImageData(imageBytes) : null;
        ImageContentType = imageContentType;
        SourceCrop = sourceCrop;
        _imageX = imageX;
        _imageY = imageY;
        _imageWidth = imageWidth;
        _imageHeight = imageHeight;
        AlternativeText = alternativeText;
    }

    private HtmlRenderDrawing(
        OfficeDrawing drawing,
        double x,
        double y,
        double width,
        double height,
        int paintOrder,
        string? alternativeText,
        string? linkUri,
        string? source,
        double layoutY,
        bool clone,
        HtmlRenderImageData? imageData,
        string? imageContentType,
        OfficeImageSourceCrop sourceCrop,
        double? imageX,
        double? imageY,
        double? imageWidth,
        double? imageHeight)
        : base(HtmlRenderVisualKind.Drawing, x, y, width, height, paintOrder, linkUri, source, layoutY) {
        _drawing = clone ? drawing.Clone() : drawing;
        _imageData = imageData;
        ImageContentType = imageContentType;
        SourceCrop = sourceCrop;
        _imageX = imageX;
        _imageY = imageY;
        _imageWidth = imageWidth;
        _imageHeight = imageHeight;
        AlternativeText = alternativeText;
    }

    /// <summary>Detached snapshot of the vector scene.</summary>
    public OfficeDrawing Drawing => _drawing.Clone();

    /// <summary>Optional alternative text inherited from the source image.</summary>
    public string? AlternativeText { get; }

    internal byte[]? ImageBytes => _imageData?.EncodedBytes;

    internal string? ImageContentType { get; }

    internal OfficeImageSourceCrop SourceCrop { get; }

    internal double ImageX => _imageX ?? X;

    internal double ImageY => _imageY ?? Y;

    internal double ImageWidth => _imageWidth ?? Width;

    internal double ImageHeight => _imageHeight ?? Height;

    internal static HtmlRenderDrawing CreateShared(
        OfficeDrawing drawing,
        double x,
        double y,
        double width,
        double height,
        int paintOrder,
        string? alternativeText,
        string? linkUri,
        string? source) =>
        new HtmlRenderDrawing(drawing, x, y, width, height, paintOrder, alternativeText, linkUri, source, y,
            clone: false, imageData: null, imageContentType: null, sourceCrop: default,
            imageX: null, imageY: null, imageWidth: null, imageHeight: null);

    internal OfficeDrawing InnerDrawing => _drawing;

    internal override HtmlRenderVisual Translate(double offsetX, double offsetY, int paintOrder) =>
        new HtmlRenderDrawing(_drawing, X + offsetX, Y + offsetY, Width, Height, paintOrder, AlternativeText,
            LinkUri, Source, LayoutY + offsetY, clone: false, _imageData, ImageContentType, SourceCrop,
            _imageX + offsetX, _imageY + offsetY, _imageWidth, _imageHeight);

    internal override HtmlRenderVisual TranslatePaint(double offsetX, double offsetY, int paintOrder) =>
        new HtmlRenderDrawing(_drawing, X + offsetX, Y + offsetY, Width, Height, paintOrder, AlternativeText,
            LinkUri, Source, LayoutY, clone: false, _imageData, ImageContentType, SourceCrop,
            _imageX + offsetX, _imageY + offsetY, _imageWidth, _imageHeight);
}