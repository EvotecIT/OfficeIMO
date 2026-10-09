using System.Collections.ObjectModel;

namespace OfficeIMO.Html;

/// <summary>
/// Ordered child visuals constrained by one or both axes of a rectangular clip.
/// Child coordinates remain in the same surface coordinate system as the group.
/// </summary>
public sealed class HtmlRenderClipGroup : HtmlRenderVisual {
    private readonly ReadOnlyCollection<HtmlRenderVisual> _visuals;

    internal HtmlRenderClipGroup(
        double x,
        double y,
        double width,
        double height,
        bool clipHorizontal,
        bool clipVertical,
        IEnumerable<HtmlRenderVisual> visuals,
        int paintOrder,
        string? source = null,
        double? layoutY = null,
        bool isViewportOverflow = false,
        bool isFlowFragment = false,
        int? legacyClipOwnerNodeId = null)
        : this(
            CreateState(x, y, width, height, clipHorizontal, clipVertical, visuals),
            x,
            y,
            width,
            height,
            clipHorizontal,
            clipVertical,
            paintOrder,
            source,
            layoutY,
            isViewportOverflow, isFlowFragment, legacyClipOwnerNodeId) {
    }

    private HtmlRenderClipGroup(
        ClipGroupState state,
        double clipX,
        double clipY,
        double clipWidth,
        double clipHeight,
        bool clipHorizontal,
        bool clipVertical,
        int paintOrder,
        string? source,
        double? layoutY,
        bool isViewportOverflow,
        bool isFlowFragment,
        int? legacyClipOwnerNodeId)
        : base(HtmlRenderVisualKind.ClipGroup, state.X, state.Y, state.Width, state.Height, paintOrder, null, source, layoutY ?? state.Y) {
        if (!clipHorizontal && !clipVertical) {
            throw new ArgumentException("A clipped render group must constrain at least one axis.", nameof(clipHorizontal));
        }
        ClipX = clipX;
        ClipY = clipY;
        ClipWidth = clipWidth;
        ClipHeight = clipHeight;
        ClipHorizontal = clipHorizontal;
        ClipVertical = clipVertical;
        IsViewportOverflow = isViewportOverflow;
        IsFlowFragment = isFlowFragment;
        LegacyClipOwnerNodeId = legacyClipOwnerNodeId;
        _visuals = state.Visuals;
    }

    /// <summary>Horizontal origin of the authored overflow clip rectangle.</summary>
    public double ClipX { get; }

    /// <summary>Vertical origin of the authored overflow clip rectangle.</summary>
    public double ClipY { get; }

    /// <summary>Width of the authored overflow clip rectangle.</summary>
    public double ClipWidth { get; }

    /// <summary>Height of the authored overflow clip rectangle.</summary>
    public double ClipHeight { get; }

    /// <summary>Whether the group clips content outside its horizontal bounds.</summary>
    public bool ClipHorizontal { get; }

    /// <summary>Whether the group clips content outside its vertical bounds.</summary>
    public bool ClipVertical { get; }

    // Root overflow clips the selected viewport after layout. It does not suppress
    // the paginated overflow that establishes automatic print fitting.
    internal bool IsViewportOverflow { get; }

    // Pagination created this clip from normal flow. Relative paint must move
    // that segment clip; authored overflow clips have already moved with paint.
    internal bool IsFlowFragment { get; }

    // Stable source identity lets extracted fixed descendants retain this paint
    // boundary after translation and page projection without retaining the DOM.
    internal int? LegacyClipOwnerNodeId { get; }

    /// <summary>Ordered child visuals in the same coordinate space as the group.</summary>
    public IReadOnlyList<HtmlRenderVisual> Visuals => _visuals;

    internal override HtmlRenderVisual TranslateCore(double offsetX, double offsetY, int paintOrder) =>
        new HtmlRenderClipGroup(
            ClipX + offsetX,
            ClipY + offsetY,
            ClipWidth,
            ClipHeight,
            ClipHorizontal,
            ClipVertical,
            _visuals.Select((visual, index) => visual.Translate(offsetX, offsetY, index)),
            paintOrder,
            Source,
            LayoutY + offsetY,
            IsViewportOverflow, IsFlowFragment, LegacyClipOwnerNodeId);

    internal override HtmlRenderVisual TranslatePaintCore(double offsetX, double offsetY, int paintOrder) =>
        new HtmlRenderClipGroup(
            ClipX + offsetX,
            ClipY + offsetY,
            ClipWidth,
            ClipHeight,
            ClipHorizontal,
            ClipVertical,
            _visuals.Select((visual, index) => visual.TranslatePaint(offsetX, offsetY, index)),
            paintOrder,
            Source,
            LayoutY,
            IsViewportOverflow, IsFlowFragment, LegacyClipOwnerNodeId);

    internal HtmlRenderVisual ProjectPaint(IEnumerable<HtmlRenderVisual> visuals, double offsetX, double offsetY, int paintOrder,
        double fragmentPaintOffsetY = 0D) =>
        new HtmlRenderClipGroup(
            ClipX + offsetX, ClipY + offsetY + fragmentPaintOffsetY, ClipWidth, ClipHeight,
            ClipHorizontal, ClipVertical, visuals, paintOrder, Source, LayoutY, IsViewportOverflow, legacyClipOwnerNodeId: LegacyClipOwnerNodeId);

    private static ClipGroupState CreateState(
        double clipX,
        double clipY,
        double clipWidth,
        double clipHeight,
        bool clipHorizontal,
        bool clipVertical,
        IEnumerable<HtmlRenderVisual> visuals) {
        var ordered = new List<HtmlRenderVisual>(visuals ?? throw new ArgumentNullException(nameof(visuals)))
            .OrderBy(item => item.PaintOrder)
            .ToList()
            .AsReadOnly();
        double left = clipHorizontal ? clipX : Math.Min(clipX, ordered.Select(item => item.X).DefaultIfEmpty(clipX).Min());
        double top = clipVertical ? clipY : Math.Min(clipY, ordered.Select(item => item.Y).DefaultIfEmpty(clipY).Min());
        double right = clipHorizontal
            ? clipX + clipWidth
            : Math.Max(clipX + clipWidth, ordered.Select(item => item.X + item.Width).DefaultIfEmpty(clipX + clipWidth).Max());
        double bottom = clipVertical
            ? clipY + clipHeight
            : Math.Max(clipY + clipHeight, ordered.Select(item => item.Y + item.Height).DefaultIfEmpty(clipY + clipHeight).Max());
        return new ClipGroupState(left, top, Math.Max(0.01D, right - left), Math.Max(0.01D, bottom - top), ordered);
    }

    private sealed class ClipGroupState {
        internal ClipGroupState(double x, double y, double width, double height, ReadOnlyCollection<HtmlRenderVisual> visuals) {
            X = x;
            Y = y;
            Width = width;
            Height = height;
            Visuals = visuals;
        }
        internal double X { get; }
        internal double Y { get; }
        internal double Width { get; }
        internal double Height { get; }
        internal ReadOnlyCollection<HtmlRenderVisual> Visuals { get; }
    }
}
