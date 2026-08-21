using System.Collections.ObjectModel;
using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

/// <summary>Native-editable layout region kind retained by the shared HTML layout owner.</summary>
public enum HtmlRenderLayoutRegionKind {
    /// <summary>Absolutely or fixed-positioned region.</summary>
    Positioned,
    /// <summary>Left or right floating region.</summary>
    Floating,
    /// <summary>Flex formatting region.</summary>
    Flex,
    /// <summary>Grid formatting region.</summary>
    Grid
}

/// <summary>
/// Paint-neutral region retaining exact rendered geometry plus editable source content for DOCX, RTF, XLSX,
/// and PPTX adapters. Child visuals remain the authoritative visual-fidelity representation.
/// </summary>
public sealed class HtmlRenderLayoutRegion : HtmlRenderVisual {
    private readonly ReadOnlyCollection<HtmlRenderVisual> _visuals;

    internal HtmlRenderLayoutRegion(
        string sourceKey,
        HtmlRenderLayoutRegionKind regionKind,
        string sourceText,
        string position,
        string floatSide,
        int zIndex,
        int backgroundLayerCount,
        int boxShadowLayerCount,
        OfficeColor? backgroundColor,
        double x,
        double y,
        double width,
        double height,
        IEnumerable<HtmlRenderVisual> visuals,
        int paintOrder,
        string? source,
        double? layoutY = null)
        : base(HtmlRenderVisualKind.LayoutRegion, x, y, width, height, paintOrder, null, source, layoutY) {
        SourceKey = sourceKey ?? throw new ArgumentNullException(nameof(sourceKey));
        RegionKind = regionKind;
        SourceText = sourceText ?? string.Empty;
        Position = string.IsNullOrWhiteSpace(position) ? "static" : position;
        FloatSide = string.IsNullOrWhiteSpace(floatSide) ? "none" : floatSide;
        ZIndex = zIndex;
        BackgroundLayerCount = Math.Max(0, backgroundLayerCount);
        BoxShadowLayerCount = Math.Max(0, boxShadowLayerCount);
        BackgroundColor = backgroundColor;
        _visuals = new List<HtmlRenderVisual>(visuals ?? throw new ArgumentNullException(nameof(visuals)))
            .OrderBy(item => item.PaintOrder)
            .ToList()
            .AsReadOnly();
    }

    /// <summary>Operation-stable source key shared by adapter DOM and layout output.</summary>
    public string SourceKey { get; }
    /// <summary>Formatting model that established this region.</summary>
    public HtmlRenderLayoutRegionKind RegionKind { get; }
    /// <summary>Editable visible source text retained for native text containers.</summary>
    public string SourceText { get; }
    /// <summary>Resolved CSS position mode.</summary>
    public string Position { get; }
    /// <summary>Resolved float side.</summary>
    public string FloatSide { get; }
    /// <summary>Resolved stacking index, with auto represented as zero.</summary>
    public int ZIndex { get; }
    /// <summary>Declared supported CSS background image layer count.</summary>
    public int BackgroundLayerCount { get; }
    /// <summary>Declared supported CSS box-shadow layer count.</summary>
    public int BoxShadowLayerCount { get; }
    /// <summary>Resolved solid background color, when present.</summary>
    public OfficeColor? BackgroundColor { get; }
    /// <summary>One-based render surface containing this unfragmented region.</summary>
    public int SurfaceNumber { get; internal set; } = 1;
    /// <summary>One-based generic semantic section that owns this region.</summary>
    public int SemanticSectionNumber { get; internal set; } = 1;
    /// <summary>Rendered horizontal origin of the owning semantic section.</summary>
    internal double SemanticSectionOriginX { get; set; }
    /// <summary>Rendered vertical origin of the owning semantic section.</summary>
    internal double SemanticSectionOriginY { get; set; }
    /// <summary>One-based owning root semantic table, or zero for narrative content.</summary>
    public int SemanticTableNumber { get; internal set; }
    /// <summary>Rendered horizontal origin of the owning root semantic table.</summary>
    internal double SemanticTableOriginX { get; set; }
    /// <summary>Rendered vertical origin of the owning root semantic table.</summary>
    internal double SemanticTableOriginY { get; set; }
    /// <summary>Ordered visual children for destination-specific native projection or fidelity fallback.</summary>
    public IReadOnlyList<HtmlRenderVisual> Visuals => _visuals;

    internal override HtmlRenderVisual Translate(double offsetX, double offsetY, int paintOrder) {
        var translated = new HtmlRenderLayoutRegion(SourceKey, RegionKind, SourceText, Position, FloatSide, ZIndex,
            BackgroundLayerCount, BoxShadowLayerCount, BackgroundColor,
            X + offsetX, Y + offsetY, Width, Height,
            _visuals.Select((visual, index) => visual.Translate(offsetX, offsetY, index)),
            paintOrder, Source, LayoutY + offsetY);
        translated.SurfaceNumber = SurfaceNumber;
        translated.SemanticSectionNumber = SemanticSectionNumber;
        translated.SemanticSectionOriginX = SemanticSectionOriginX + offsetX;
        translated.SemanticSectionOriginY = SemanticSectionOriginY + offsetY;
        translated.SemanticTableNumber = SemanticTableNumber;
        translated.SemanticTableOriginX = SemanticTableOriginX + offsetX;
        translated.SemanticTableOriginY = SemanticTableOriginY + offsetY;
        return translated;
    }

    internal override HtmlRenderVisual TranslatePaint(double offsetX, double offsetY, int paintOrder) {
        var translated = new HtmlRenderLayoutRegion(SourceKey, RegionKind, SourceText, Position, FloatSide, ZIndex,
            BackgroundLayerCount, BoxShadowLayerCount, BackgroundColor,
            X + offsetX, Y + offsetY, Width, Height,
            _visuals.Select((visual, index) => visual.TranslatePaint(offsetX, offsetY, index)),
            paintOrder, Source, LayoutY);
        translated.SurfaceNumber = SurfaceNumber;
        translated.SemanticSectionNumber = SemanticSectionNumber;
        translated.SemanticSectionOriginX = SemanticSectionOriginX + offsetX;
        translated.SemanticSectionOriginY = SemanticSectionOriginY;
        translated.SemanticTableNumber = SemanticTableNumber;
        translated.SemanticTableOriginX = SemanticTableOriginX + offsetX;
        translated.SemanticTableOriginY = SemanticTableOriginY;
        return translated;
    }
}
