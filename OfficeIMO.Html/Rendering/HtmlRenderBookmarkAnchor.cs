namespace OfficeIMO.Html;

/// <summary>Paint-neutral navigation anchor retained for elements without searchable text visuals.</summary>
public sealed class HtmlRenderBookmarkAnchor : HtmlRenderVisual {
    internal HtmlRenderBookmarkAnchor(
        int semanticNodeId,
        string text,
        double x,
        double y,
        double width,
        double height,
        int paintOrder,
        string? source,
        double? layoutY = null)
        : base(HtmlRenderVisualKind.BookmarkAnchor, x, y, width, height, paintOrder, null, source, layoutY) {
        SemanticNodeId = semanticNodeId;
        Text = text ?? throw new ArgumentNullException(nameof(text));
    }

    /// <summary>Stable operation-scoped semantic node identifier.</summary>
    public int SemanticNodeId { get; }

    /// <summary>Rendered fallback label used when CSS did not provide <c>bookmark-label</c>.</summary>
    public string Text { get; }

    internal override HtmlRenderVisual Translate(double offsetX, double offsetY, int paintOrder) =>
        new HtmlRenderBookmarkAnchor(SemanticNodeId, Text, X + offsetX, Y + offsetY, Width, Height, paintOrder, Source, LayoutY + offsetY);

    internal override HtmlRenderVisual TranslatePaint(double offsetX, double offsetY, int paintOrder) =>
        new HtmlRenderBookmarkAnchor(SemanticNodeId, Text, X + offsetX, Y + offsetY, Width, Height, paintOrder, Source, LayoutY);
}
