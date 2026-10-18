using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

/// <summary>
/// Positioned vector shape produced by HTML paint preparation.
/// </summary>
public sealed class HtmlRenderShape : HtmlRenderVisual {
    private readonly OfficeShape _shape;

    internal HtmlRenderShape(OfficeShape shape, double x, double y, int paintOrder, string? linkUri = null, string? source = null, double? layoutY = null, double? layoutHeight = null, bool isAtomicReplacedPlaceholder = false, bool canExtendFragmentBackground = false, bool isCollapsedTableBorder = false, double? collapsedBorderBoundaryY = null, HtmlRenderBorderEdges? fragmentSideBorders = null, bool canExtendFragmentColumnRule = false)
        : base(HtmlRenderVisualKind.Shape, x, y, shape?.Width ?? 0D, shape?.Height ?? 0D, paintOrder, linkUri, source, layoutY, layoutHeight) {
        _shape = shape?.Clone() ?? throw new ArgumentNullException(nameof(shape));
        IsAtomicReplacedPlaceholder = isAtomicReplacedPlaceholder;
        CanExtendFragmentBackground = canExtendFragmentBackground;
        FragmentSideBorders = fragmentSideBorders;
        IsCollapsedTableBorder = isCollapsedTableBorder;
        CollapsedBorderBoundaryY = collapsedBorderBoundaryY;
        CanExtendFragmentColumnRule = canExtendFragmentColumnRule;
    }

    internal bool IsAtomicReplacedPlaceholder { get; }

    /// <summary>Plain auto-height box paint that may cover unused space in a continuing page fragment.</summary>
    internal bool CanExtendFragmentBackground { get; }

    /// <summary>Rule paint before a deferred spanner may cover the current page's unused body space.</summary>
    internal bool CanExtendFragmentColumnRule { get; }

    /// <summary>Own solid box sides accompanying an eligible slice background through fragment slack.</summary>
    internal HtmlRenderBorderEdges? FragmentSideBorders { get; }

    /// <summary>Grid paint that follows cell backgrounds and retains full strokes at row-boundary cuts.</summary>
    internal bool IsCollapsedTableBorder { get; }

    /// <summary>Logical horizontal grid edge shared by all strokes of a collapsed rule.</summary>
    internal double? CollapsedBorderBoundaryY { get; }

    /// <summary>Independent snapshot of the shared dependency-free vector shape.</summary>
    public OfficeShape Shape => _shape.Clone();

    internal OfficeShape InnerShape => _shape;

    internal override HtmlRenderVisual TranslateCore(double offsetX, double offsetY, int paintOrder) =>
        new HtmlRenderShape(_shape, X + offsetX, Y + offsetY, paintOrder, LinkUri, Source, LayoutY + offsetY, LayoutHeight, IsAtomicReplacedPlaceholder, CanExtendFragmentBackground, IsCollapsedTableBorder, CollapsedBorderBoundaryY + offsetY, FragmentSideBorders, CanExtendFragmentColumnRule);

    internal override HtmlRenderVisual TranslatePaintCore(double offsetX, double offsetY, int paintOrder) =>
        new HtmlRenderShape(_shape, X + offsetX, Y + offsetY, paintOrder, LinkUri, Source, LayoutY, LayoutHeight, IsAtomicReplacedPlaceholder, CanExtendFragmentBackground, IsCollapsedTableBorder, CollapsedBorderBoundaryY, FragmentSideBorders, CanExtendFragmentColumnRule);
}
