using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

/// <summary>Positioned shared vector drawing emitted by HTML layout.</summary>
public sealed class HtmlRenderDrawing : HtmlRenderVisual {
    private readonly OfficeDrawing _drawing;

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
        double? layoutY = null)
        : base(HtmlRenderVisualKind.Drawing, x, y, width, height, paintOrder, linkUri, source, layoutY) {
        _drawing = (drawing ?? throw new ArgumentNullException(nameof(drawing))).Clone();
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
        bool clone)
        : base(HtmlRenderVisualKind.Drawing, x, y, width, height, paintOrder, linkUri, source, layoutY) {
        _drawing = clone ? drawing.Clone() : drawing;
        AlternativeText = alternativeText;
    }

    /// <summary>Detached snapshot of the vector scene.</summary>
    public OfficeDrawing Drawing => _drawing.Clone();

    /// <summary>Optional alternative text inherited from the source image.</summary>
    public string? AlternativeText { get; }

    internal OfficeDrawing InnerDrawing => _drawing;

    internal override HtmlRenderVisual Translate(double offsetX, double offsetY, int paintOrder) =>
        new HtmlRenderDrawing(_drawing, X + offsetX, Y + offsetY, Width, Height, paintOrder, AlternativeText, LinkUri, Source, LayoutY + offsetY, clone: false);

    internal override HtmlRenderVisual TranslatePaint(double offsetX, double offsetY, int paintOrder) =>
        new HtmlRenderDrawing(_drawing, X + offsetX, Y + offsetY, Width, Height, paintOrder, AlternativeText, LinkUri, Source, LayoutY, clone: false);
}
