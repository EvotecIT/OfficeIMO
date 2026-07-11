using System.Collections.ObjectModel;

namespace OfficeIMO.Html;

/// <summary>Positioned paint fragments that represent one logical text value for extraction.</summary>
public sealed class HtmlRenderLogicalTextGroup : HtmlRenderVisual {
    private readonly ReadOnlyCollection<HtmlRenderVisual> _visuals;

    internal HtmlRenderLogicalTextGroup(
        string text,
        double x,
        double y,
        double width,
        double height,
        IEnumerable<HtmlRenderVisual> visuals,
        int paintOrder,
        string? source,
        double? layoutY = null)
        : base(HtmlRenderVisualKind.LogicalTextGroup, x, y, width, height, paintOrder, null, source, layoutY) {
        Text = text ?? throw new ArgumentNullException(nameof(text));
        _visuals = new List<HtmlRenderVisual>(visuals ?? throw new ArgumentNullException(nameof(visuals)))
            .OrderBy(item => item.PaintOrder)
            .ToList()
            .AsReadOnly();
    }

    /// <summary>Logical source-order text represented by the positioned child fragments.</summary>
    public string Text { get; }

    /// <summary>Ordered positioned paint fragments.</summary>
    public IReadOnlyList<HtmlRenderVisual> Visuals => _visuals;

    internal override HtmlRenderVisual Translate(double offsetX, double offsetY, int paintOrder) =>
        new HtmlRenderLogicalTextGroup(Text, X + offsetX, Y + offsetY, Width, Height, _visuals.Select((visual, index) => visual.Translate(offsetX, offsetY, index)), paintOrder, Source, LayoutY + offsetY);

    internal override HtmlRenderVisual TranslatePaint(double offsetX, double offsetY, int paintOrder) =>
        new HtmlRenderLogicalTextGroup(Text, X + offsetX, Y + offsetY, Width, Height, _visuals.Select((visual, index) => visual.TranslatePaint(offsetX, offsetY, index)), paintOrder, Source, LayoutY);
}
