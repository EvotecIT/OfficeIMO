using System.Collections.ObjectModel;
using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

/// <summary>
/// Backend-neutral ordered paint group with one affine transform and isolated opacity.
/// </summary>
public sealed class HtmlRenderEffectGroup : HtmlRenderVisual {
    private readonly ReadOnlyCollection<HtmlRenderVisual> _visuals;

    internal HtmlRenderEffectGroup(
        double x,
        double y,
        double width,
        double height,
        OfficeTransform transform,
        double opacity,
        IEnumerable<HtmlRenderVisual> visuals,
        int paintOrder,
        string? source,
        double? layoutY = null)
        : base(HtmlRenderVisualKind.EffectGroup, x, y, width, height, paintOrder, null, source, layoutY) {
        if (double.IsNaN(opacity) || double.IsInfinity(opacity) || opacity < 0D || opacity > 1D) {
            throw new ArgumentOutOfRangeException(nameof(opacity), "Effect-group opacity must be between zero and one.");
        }
        Transform = transform;
        Opacity = opacity;
        _visuals = new List<HtmlRenderVisual>(visuals ?? throw new ArgumentNullException(nameof(visuals)))
            .OrderBy(item => item.PaintOrder)
            .ToList()
            .AsReadOnly();
    }

    /// <summary>Destination-space affine transform applied to the group.</summary>
    public OfficeTransform Transform { get; }

    /// <summary>Isolated group opacity from zero through one.</summary>
    public double Opacity { get; }

    /// <summary>Ordered child visuals.</summary>
    public IReadOnlyList<HtmlRenderVisual> Visuals => _visuals;

    internal override HtmlRenderVisual Translate(double offsetX, double offsetY, int paintOrder) {
        OfficeTransform transform = RebaseTransform(Transform, offsetX, offsetY);
        return new HtmlRenderEffectGroup(
            X + offsetX,
            Y + offsetY,
            Width,
            Height,
            transform,
            Opacity,
            _visuals.Select((visual, index) => visual.Translate(offsetX, offsetY, index)),
            paintOrder,
            Source,
            LayoutY + offsetY);
    }

    internal override HtmlRenderVisual TranslatePaint(double offsetX, double offsetY, int paintOrder) {
        OfficeTransform transform = RebaseTransform(Transform, offsetX, offsetY);
        return new HtmlRenderEffectGroup(
            X + offsetX,
            Y + offsetY,
            Width,
            Height,
            transform,
            Opacity,
            _visuals.Select((visual, index) => visual.TranslatePaint(offsetX, offsetY, index)),
            paintOrder,
            Source,
            LayoutY);
    }

    private static OfficeTransform RebaseTransform(OfficeTransform transform, double offsetX, double offsetY) =>
        OfficeTransform.Translate(-offsetX, -offsetY)
            .Then(transform)
            .Then(OfficeTransform.Translate(offsetX, offsetY));
}
