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
        double? layoutY = null,
        double? layoutHeight = null,
        HtmlRenderLogicalTextScope? logicalScope = null)
        : base(HtmlRenderVisualKind.LogicalTextGroup, x, y, width, height, paintOrder, null, source, layoutY, layoutHeight) {
        LogicalScope = logicalScope;
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

    internal HtmlRenderLogicalTextScope? LogicalScope { get; }

    internal override HtmlRenderVisual TranslateCore(double offsetX, double offsetY, int paintOrder) =>
        new HtmlRenderLogicalTextGroup(Text, X + offsetX, Y + offsetY, Width, Height, _visuals.Select((visual, index) => visual.Translate(offsetX, offsetY, index)), paintOrder, Source, LayoutY + offsetY, LayoutHeight, LogicalScope);

    internal override HtmlRenderVisual TranslatePaintCore(double offsetX, double offsetY, int paintOrder) =>
        new HtmlRenderLogicalTextGroup(Text, X + offsetX, Y + offsetY, Width, Height, _visuals.Select((visual, index) => visual.TranslatePaint(offsetX, offsetY, index)), paintOrder, Source, LayoutY, LayoutHeight, LogicalScope);

    internal HtmlRenderVisual ProjectPaint(IEnumerable<HtmlRenderVisual> visuals, double offsetX, double offsetY, int paintOrder, bool ownsLogicalText) =>
        new HtmlRenderLogicalTextGroup(ownsLogicalText ? Text : string.Empty,
            X + offsetX, Y + offsetY, Width, Height, visuals, paintOrder, Source, LayoutY, LayoutHeight, LogicalScope);
}

/// <summary>Shared identity of discontiguous paint fragments for one logical text scope.</summary>
internal sealed class HtmlRenderLogicalTextScope { }
