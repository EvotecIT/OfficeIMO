using System.Collections.ObjectModel;
using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

/// <summary>
/// Ordered child visuals clipped by one shared Drawing path in surface coordinates.
/// Child coordinates remain in the same surface coordinate system as the group.
/// </summary>
public sealed class HtmlRenderPathClipGroup : HtmlRenderVisual {
    private readonly ReadOnlyCollection<HtmlRenderVisual> _visuals;

    internal HtmlRenderPathClipGroup(
        double x,
        double y,
        OfficeClipPath clipPath,
        IEnumerable<HtmlRenderVisual> visuals,
        int paintOrder,
        string? source = null,
        double? layoutY = null)
        : base(
            HtmlRenderVisualKind.PathClipGroup,
            x,
            y,
            clipPath?.Width ?? 0D,
            clipPath?.Height ?? 0D,
            paintOrder,
            null,
            source,
            layoutY) {
        ClipPath = clipPath?.Clone() ?? throw new ArgumentNullException(nameof(clipPath));
        _visuals = new List<HtmlRenderVisual>(visuals ?? throw new ArgumentNullException(nameof(visuals)))
            .OrderBy(item => item.PaintOrder)
            .ToList()
            .AsReadOnly();
    }

    /// <summary>Horizontal origin of the clip path in surface coordinates.</summary>
    public double ClipX => X;

    /// <summary>Vertical origin of the clip path in surface coordinates.</summary>
    public double ClipY => Y;

    /// <summary>Detached shared Drawing clip geometry in local group coordinates.</summary>
    public OfficeClipPath ClipPath { get; }

    /// <summary>Ordered child visuals in the same surface coordinate system as the group.</summary>
    public IReadOnlyList<HtmlRenderVisual> Visuals => _visuals;

    internal override HtmlRenderVisual Translate(double offsetX, double offsetY, int paintOrder) =>
        new HtmlRenderPathClipGroup(
            ClipX + offsetX,
            ClipY + offsetY,
            ClipPath,
            _visuals.Select((visual, index) => visual.Translate(offsetX, offsetY, index)),
            paintOrder,
            Source,
            LayoutY + offsetY);

    internal override HtmlRenderVisual TranslatePaint(double offsetX, double offsetY, int paintOrder) =>
        new HtmlRenderPathClipGroup(
            ClipX + offsetX,
            ClipY + offsetY,
            ClipPath,
            _visuals.Select((visual, index) => visual.TranslatePaint(offsetX, offsetY, index)),
            paintOrder,
            Source,
            LayoutY);
}
