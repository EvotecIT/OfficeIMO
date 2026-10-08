using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

/// <summary>
/// Positioned vector shape produced by HTML paint preparation.
/// </summary>
public sealed class HtmlRenderShape : HtmlRenderVisual {
    private readonly OfficeShape _shape;

    internal HtmlRenderShape(OfficeShape shape, double x, double y, int paintOrder, string? linkUri = null, string? source = null, double? layoutY = null, double? layoutHeight = null, bool isAtomicReplacedPlaceholder = false, bool canExtendFragmentBackground = false)
        : base(HtmlRenderVisualKind.Shape, x, y, shape?.Width ?? 0D, shape?.Height ?? 0D, paintOrder, linkUri, source, layoutY, layoutHeight) {
        _shape = shape?.Clone() ?? throw new ArgumentNullException(nameof(shape));
        IsAtomicReplacedPlaceholder = isAtomicReplacedPlaceholder;
        CanExtendFragmentBackground = canExtendFragmentBackground;
    }

    internal bool IsAtomicReplacedPlaceholder { get; }

    /// <summary>Plain auto-height box paint that may cover unused space in a continuing page fragment.</summary>
    internal bool CanExtendFragmentBackground { get; }

    /// <summary>Independent snapshot of the shared dependency-free vector shape.</summary>
    public OfficeShape Shape => _shape.Clone();

    internal OfficeShape InnerShape => _shape;

    internal override HtmlRenderVisual TranslateCore(double offsetX, double offsetY, int paintOrder) =>
        new HtmlRenderShape(_shape, X + offsetX, Y + offsetY, paintOrder, LinkUri, Source, LayoutY + offsetY, LayoutHeight, IsAtomicReplacedPlaceholder, CanExtendFragmentBackground);

    internal override HtmlRenderVisual TranslatePaintCore(double offsetX, double offsetY, int paintOrder) =>
        new HtmlRenderShape(_shape, X + offsetX, Y + offsetY, paintOrder, LinkUri, Source, LayoutY, LayoutHeight, IsAtomicReplacedPlaceholder, CanExtendFragmentBackground);
}
