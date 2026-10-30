using OfficeIMO.Drawing;
using OfficeIMO.Drawing.Binary;

namespace OfficeIMO.Publisher.Internal;

/// <summary>Maps native OfficeArt geometry and paint once for both semantic import and retained native artwork projection.</summary>
internal sealed partial class PublisherNativeShapeMapper {
    private readonly IReadOnlyList<OfficeColor?> _palette;
    private readonly PublisherParseContext _context;
    internal PublisherNativeShapeMapper(IReadOnlyList<OfficeColor?> palette, PublisherParseContext context) {
        _palette = palette; _context = context;
    }

    /// <summary>Preserves logical native extents for custom paths and fills; only preset fallback geometry uses the positive render canvas. Native lines always retain logical extents.</summary>
    internal OfficeShape Map(PublisherEscherShape source, double geometryWidth, double geometryHeight,
        double canvasWidth, double canvasHeight, out OfficeShape? clipShape) {
        OfficeArtCustomPathProjection? custom = CustomPath(source, geometryWidth, geometryHeight);
        OfficeShape shape = custom?.Shape ?? Geometry(source, source.Type == 20 ? geometryWidth : canvasWidth,
            source.Type == 20 ? geometryHeight : canvasHeight);
        OfficeArtShapeStyle style = source.Style;
        bool noFill = custom?.NoFill ?? OfficeArtCustomPathProjector.IsFillDisabled(source.Properties);
        bool noLine = custom?.NoLine ?? OfficeArtCustomPathProjector.IsLineDisabled(source.Properties);
        shape.FillColor = style.FillEnabled != false && !noFill ? Resolve(style.FillColor, source.Id) : null;
        shape.StrokeColor = style.LineEnabled != false && !noLine ? Resolve(style.LineColor, source.Id) : null;
        shape.StrokeWidth = Math.Max(0, (style.LineWidthEmus ?? 9525) / 12700D);
        shape.FillOpacity = style.FillOpacity; shape.StrokeOpacity = style.LineOpacity;
        if (!noFill) ProjectFill(source, shape, geometryWidth, geometryHeight);
        ProjectLineDetails(source, shape);
        if (style.HasProjectableShadow) {
            shape.Shadow = new OfficeShadow(Resolve(style.ShadowColor, source.Id) ?? OfficeColor.Black, style.ShadowOpacity ?? 1,
                (style.ShadowOffsetXEmus ?? 0) / 12700D, (style.ShadowOffsetYEmus ?? 0) / 12700D, Math.Max(0, (style.ShadowSoftnessEmus ?? 0) / 12700D));
        }
        clipShape = custom?.Shape;
        return shape;
    }
    private OfficeShape Geometry(PublisherEscherShape source, double width, double height) {
        switch (source.Type) {
            case 1: case 75: case 202: return OfficeShape.Rectangle(width, height);
            case 2: return OfficeShape.RoundedRectangle(width, height, Math.Min(width, height) / 6);
            case 3: return OfficeShape.Ellipse(width, height);
            case 4: return OfficeShape.Polygon(new OfficePoint(width / 2, 0), new OfficePoint(width, height / 2), new OfficePoint(width / 2, height), new OfficePoint(0, height / 2));
            case 5: return OfficeShape.Polygon(new OfficePoint(width / 2, 0), new OfficePoint(width, height), new OfficePoint(0, height));
            case 6: return OfficeShape.Polygon(new OfficePoint(0, 0), new OfficePoint(width, height), new OfficePoint(0, height));
            case 20: return OfficeShape.Line(0, 0, width, height);
            default:
                _context.Add("PUB_SHAPE_GEOMETRY_APPROXIMATED", $"Native shape type {source.Type} was approximated by its bounding rectangle.", OfficeConversionLossKind.Approximation, PublisherEscherReader.ShapeLocation(source.Id));
                return OfficeShape.Rectangle(width, height);
        }
    }
    private OfficeColor? Resolve(OfficeArtColorReference? reference, uint id) {
        if (!reference.HasValue || reference.Value.IsIgnored) return null;
        if (reference.Value.TryResolve(index => index < _palette.Count ? _palette[index] : null, out OfficeColor color)) return color;
        _context.Add("PUB_SHAPE_COLOR_UNRESOLVED", "A native color reference was approximated as black.", OfficeConversionLossKind.Approximation, PublisherEscherReader.ShapeLocation(id));
        return OfficeColor.Black;
    }
}
