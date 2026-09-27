namespace OfficeIMO.Html;

/// <summary>
/// Temporary border-box geometry used to measure paginated scrollable overflow.
/// It follows the same translation, clipping and effect groups as paint, but is
/// removed before the retained document is exposed to any backend or caller.
/// </summary>
internal sealed class HtmlRenderLayoutBox : HtmlRenderVisual {
    internal HtmlRenderLayoutBox(double x, double y, double width, double height,
        int paintOrder, string? source, double? layoutY = null)
        : base(HtmlRenderVisualKind.Shape, x, y, width, height, paintOrder, null, source, layoutY) { }

    internal override HtmlRenderVisual Translate(double offsetX, double offsetY, int paintOrder) =>
        new HtmlRenderLayoutBox(X + offsetX, Y + offsetY, Width, Height, paintOrder, Source, LayoutY + offsetY);

    internal override HtmlRenderVisual TranslatePaint(double offsetX, double offsetY, int paintOrder) =>
        new HtmlRenderLayoutBox(X + offsetX, Y + offsetY, Width, Height, paintOrder, Source, LayoutY);
}
