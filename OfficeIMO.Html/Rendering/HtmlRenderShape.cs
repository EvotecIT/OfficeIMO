using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

/// <summary>
/// Positioned vector shape produced by HTML paint preparation.
/// </summary>
public sealed class HtmlRenderShape : HtmlRenderVisual {
    internal HtmlRenderShape(OfficeShape shape, double x, double y, int paintOrder, string? linkUri = null, string? source = null, double? layoutY = null)
        : base(HtmlRenderVisualKind.Shape, x, y, shape?.Width ?? 0D, shape?.Height ?? 0D, paintOrder, linkUri, source, layoutY) {
        Shape = shape?.Clone() ?? throw new ArgumentNullException(nameof(shape));
    }

    /// <summary>Shared dependency-free vector shape.</summary>
    public OfficeShape Shape { get; }

    internal override HtmlRenderVisual Translate(double offsetX, double offsetY, int paintOrder) =>
        new HtmlRenderShape(Shape, X + offsetX, Y + offsetY, paintOrder, LinkUri, Source, LayoutY + offsetY);

    internal override HtmlRenderVisual TranslatePaint(double offsetX, double offsetY, int paintOrder) =>
        new HtmlRenderShape(Shape, X + offsetX, Y + offsetY, paintOrder, LinkUri, Source, LayoutY);
}
