namespace OfficeIMO.Html;

/// <summary>
/// Temporary border-box geometry used to measure paginated scrollable overflow.
/// It follows the same translation, clipping and effect groups as paint, but is
/// removed before the retained document is exposed to any backend or caller.
/// </summary>
internal sealed class HtmlRenderLayoutBox : HtmlRenderVisual {
    internal HtmlRenderLayoutBox(double x, double y, double width, double height,
        int paintOrder, string? source, double? layoutY = null, bool isAbsolutePrintOverflow = false)
        : base(HtmlRenderVisualKind.Shape, x, y, width, height, paintOrder, null, source, layoutY) {
        IsAbsolutePrintOverflow = isAbsolutePrintOverflow;
    }

    internal bool IsAbsolutePrintOverflow { get; private set; }

    internal void MarkAbsolutePrintOverflow() => IsAbsolutePrintOverflow = true;

    internal override HtmlRenderVisual TranslateCore(double offsetX, double offsetY, int paintOrder) =>
        new HtmlRenderLayoutBox(X + offsetX, Y + offsetY, Width, Height, paintOrder, Source,
            LayoutY + offsetY, IsAbsolutePrintOverflow);

    internal override HtmlRenderVisual TranslatePaintCore(double offsetX, double offsetY, int paintOrder) =>
        new HtmlRenderLayoutBox(X + offsetX, Y + offsetY, Width, Height, paintOrder, Source,
            LayoutY, IsAbsolutePrintOverflow);
}
