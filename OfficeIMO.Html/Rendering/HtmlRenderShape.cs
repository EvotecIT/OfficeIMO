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

    /// <summary>Creates and retains a private line for one resolved collapsed-border stroke.</summary>
    internal static HtmlRenderShape CreateCollapsedBorderStroke(
        bool horizontal, double x, double y, double length, OfficeColor color, double width,
        OfficeStrokeDashStyle dashStyle, OfficeStrokeLineCap lineCap, int paintOrder,
        string source, double? horizontalBoundaryY = null) {
        OfficeShape shape = horizontal
            ? OfficeShape.Line(0D, 0D, length, 0.0001D)
            : OfficeShape.Line(0D, 0D, 0.0001D, length);
        shape.StrokeColor = color;
        shape.StrokeWidth = width;
        shape.StrokeDashStyle = dashStyle;
        shape.StrokeLineCap = lineCap;
        shape.FillColor = null;
        // Retain the horizontal stroke's extent when pagination slices between rows.
        double layoutHeight = horizontal ? Math.Max(0.01D, width) : shape.Height;
        double layoutY = horizontal ? y - layoutHeight / 2D : y;
        return new HtmlRenderShape(shape, x, y, paintOrder, source, layoutY, layoutHeight,
            horizontal ? horizontalBoundaryY ?? y : null);
    }

    // Only the collapsed-border factory supplies this private, unaliased line.
    private HtmlRenderShape(OfficeShape collapsedBorder, double x, double y, int paintOrder,
        string source, double layoutY, double layoutHeight, double? collapsedBorderBoundaryY)
        : base(HtmlRenderVisualKind.Shape, x, y, collapsedBorder.Width, collapsedBorder.Height,
            paintOrder, null, source, layoutY, layoutHeight) {
        _shape = collapsedBorder;
        IsCollapsedTableBorder = true;
        CollapsedBorderBoundaryY = collapsedBorderBoundaryY;
    }

    // A visual's private geometry is immutable: public Shape reads return detached
    // copies, and internal readers clone before editing. Position-only projections
    // can therefore share this snapshot without another shape/effect graph copy.
    private HtmlRenderShape(HtmlRenderShape source, double x, double y, int paintOrder,
        double layoutY, double? collapsedBorderBoundaryY)
        : base(HtmlRenderVisualKind.Shape, x, y, source.Width, source.Height, paintOrder,
            source.LinkUri, source.Source, layoutY, source.LayoutHeight) {
        _shape = source._shape;
        IsAtomicReplacedPlaceholder = source.IsAtomicReplacedPlaceholder;
        CanExtendFragmentBackground = source.CanExtendFragmentBackground;
        FragmentSideBorders = source.FragmentSideBorders;
        IsCollapsedTableBorder = source.IsCollapsedTableBorder;
        CollapsedBorderBoundaryY = collapsedBorderBoundaryY;
        CanExtendFragmentColumnRule = source.CanExtendFragmentColumnRule;
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

    /// <summary>Owned immutable geometry for internal readers; clone it before editing.</summary>
    internal OfficeShape InnerShape => _shape;

    internal override HtmlRenderVisual TranslateCore(double offsetX, double offsetY, int paintOrder) =>
        new HtmlRenderShape(this, X + offsetX, Y + offsetY, paintOrder, LayoutY + offsetY, CollapsedBorderBoundaryY + offsetY);

    internal override HtmlRenderVisual TranslatePaintCore(double offsetX, double offsetY, int paintOrder) =>
        new HtmlRenderShape(this, X + offsetX, Y + offsetY, paintOrder, LayoutY, CollapsedBorderBoundaryY);

    internal override HtmlRenderVisual TranslateThrough(HtmlRenderOffset[] offsets, int paintOrder) {
        double x = X, y = Y, layoutY = LayoutY;
        double? boundaryY = CollapsedBorderBoundaryY;
        foreach (HtmlRenderOffset offset in offsets) {
            x += offset.X;
            y += offset.Y;
            if (offset.PaintOnly) continue;
            layoutY += offset.Y;
            boundaryY += offset.Y;
        }
        return CopyStackingContextTo(new HtmlRenderShape(this, x, y, paintOrder, layoutY, boundaryY));
    }
}
