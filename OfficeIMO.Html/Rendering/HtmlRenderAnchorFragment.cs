namespace OfficeIMO.Html;

/// <summary>A paint-neutral, anchor-owned link rectangle for one laid-out fragment.</summary>
public sealed class HtmlRenderAnchorFragment : HtmlRenderVisual {
    internal HtmlRenderAnchorFragment(int anchorNodeId, string linkUri, string? linkContents,
        double x, double y, double width, double height, int paintOrder, string? source,
        double? layoutY = null)
        : base(HtmlRenderVisualKind.AnchorFragment, x, y, width, height, paintOrder,
            linkUri ?? throw new ArgumentNullException(nameof(linkUri)), source, layoutY) {
        AnchorNodeId = anchorNodeId;
        LinkContents = linkContents;
    }

    /// <summary>Identity of the source anchor, independent of its destination URL.</summary>
    public int AnchorNodeId { get; }

    /// <summary>Readable link text retained in the PDF annotation.</summary>
    public string? LinkContents { get; }

    internal override HtmlRenderVisual Translate(double offsetX, double offsetY, int paintOrder) =>
        new HtmlRenderAnchorFragment(AnchorNodeId, LinkUri!, LinkContents, X + offsetX, Y + offsetY,
            Width, Height, paintOrder, Source, LayoutY + offsetY);

    internal override HtmlRenderVisual TranslatePaint(double offsetX, double offsetY, int paintOrder) =>
        new HtmlRenderAnchorFragment(AnchorNodeId, LinkUri!, LinkContents, X + offsetX, Y + offsetY,
            Width, Height, paintOrder, Source, LayoutY);
}
