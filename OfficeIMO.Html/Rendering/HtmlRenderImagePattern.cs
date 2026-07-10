using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

/// <summary>
/// Bounded repeating image pattern retained as one backend-neutral paint operation.
/// </summary>
public sealed class HtmlRenderImagePattern : HtmlRenderVisual {
    private readonly HtmlRenderImageData _data;

    internal HtmlRenderImagePattern(
        byte[] bytes,
        string contentType,
        OfficeImagePatternLayout layout,
        int maximumTileCount,
        int paintOrder,
        string? source = null,
        double? layoutY = null)
        : this(new HtmlRenderImageData(bytes), contentType, layout, maximumTileCount, paintOrder, source, layoutY) {
    }

    private HtmlRenderImagePattern(
        HtmlRenderImageData data,
        string contentType,
        OfficeImagePatternLayout layout,
        int maximumTileCount,
        int paintOrder,
        string? source,
        double? layoutY)
        : base(
            HtmlRenderVisualKind.ImagePattern,
            layout.Area.X,
            layout.Area.Y,
            layout.Area.Width,
            layout.Area.Height,
            paintOrder,
            linkUri: null,
            source,
            layoutY) {
        if (maximumTileCount <= 0) throw new ArgumentOutOfRangeException(nameof(maximumTileCount));
        if (layout.EstimatedTileCount > maximumTileCount) throw new ArgumentException("Image pattern exceeds the configured tile-count limit.", nameof(layout));
        _data = data ?? throw new ArgumentNullException(nameof(data));
        ContentType = string.IsNullOrWhiteSpace(contentType) ? "application/octet-stream" : contentType;
        Pattern = layout;
        MaximumTileCount = maximumTileCount;
    }

    /// <summary>Detached encoded source-image bytes.</summary>
    public byte[] Bytes => _data.Snapshot();

    /// <summary>Image media type.</summary>
    public string ContentType { get; }

    /// <summary>Pattern area, origin tile, and repeat axes.</summary>
    public OfficeImagePatternLayout Pattern { get; }

    /// <summary>Maximum tile count accepted by expansion-based backends.</summary>
    public int MaximumTileCount { get; }

    internal byte[] EncodedBytes => _data.EncodedBytes;

    internal override HtmlRenderVisual Translate(double offsetX, double offsetY, int paintOrder) =>
        new HtmlRenderImagePattern(_data, ContentType, Pattern.Translate(offsetX, offsetY), MaximumTileCount, paintOrder, Source, LayoutY + offsetY);

    internal override HtmlRenderVisual TranslatePaint(double offsetX, double offsetY, int paintOrder) =>
        new HtmlRenderImagePattern(_data, ContentType, Pattern.Translate(offsetX, offsetY), MaximumTileCount, paintOrder, Source, LayoutY);
}
