namespace OfficeIMO.Html;

/// <summary>Paint-neutral named navigation destination retained by the shared render scene.</summary>
public sealed class HtmlRenderNamedDestination : HtmlRenderVisual {
    internal HtmlRenderNamedDestination(
        string name,
        double x,
        double y,
        int paintOrder,
        string? source,
        double? layoutY = null)
        : base(HtmlRenderVisualKind.NamedDestination, x, y, 0.01D, 0.01D, paintOrder, null, source, layoutY) {
        if (string.IsNullOrWhiteSpace(name)) throw new ArgumentException("Named destinations cannot be empty or whitespace.", nameof(name));
        Name = name.Trim();
    }

    /// <summary>Stable destination name used by document-internal links.</summary>
    public string Name { get; }

    internal override HtmlRenderVisual Translate(double offsetX, double offsetY, int paintOrder) =>
        new HtmlRenderNamedDestination(Name, X + offsetX, Y + offsetY, paintOrder, Source, LayoutY + offsetY);

    internal override HtmlRenderVisual TranslatePaint(double offsetX, double offsetY, int paintOrder) =>
        new HtmlRenderNamedDestination(Name, X + offsetX, Y + offsetY, paintOrder, Source, LayoutY);
}
