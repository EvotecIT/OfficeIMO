using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

/// <summary>
/// Positioned vector shape produced by HTML paint preparation.
/// </summary>
public sealed class HtmlRenderShape : HtmlRenderVisual {
    private readonly OfficeShape _shape;

    internal HtmlRenderShape(OfficeShape shape, double x, double y, int paintOrder, string? linkUri = null, string? source = null, double? layoutY = null)
        : base(HtmlRenderVisualKind.Shape, x, y, shape?.Width ?? 0D, shape?.Height ?? 0D, paintOrder, linkUri, source, layoutY) {
        _shape = shape?.Clone() ?? throw new ArgumentNullException(nameof(shape));
    }

    /// <summary>Independent snapshot of the shared dependency-free vector shape.</summary>
    public OfficeShape Shape => _shape.Clone();

    internal OfficeShape InnerShape => _shape;

    internal override HtmlRenderVisual Translate(double offsetX, double offsetY, int paintOrder) =>
        new HtmlRenderShape(_shape, X + offsetX, Y + offsetY, paintOrder, LinkUri, Source, LayoutY + offsetY);

    internal override HtmlRenderVisual TranslatePaint(double offsetX, double offsetY, int paintOrder) =>
        new HtmlRenderShape(_shape, X + offsetX, Y + offsetY, paintOrder, LinkUri, Source, LayoutY);
}
