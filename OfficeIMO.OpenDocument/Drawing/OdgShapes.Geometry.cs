using OfficeIMO.Drawing;

namespace OfficeIMO.OpenDocument;

public sealed partial class OdgShapes {
    /// <summary>Adds a rectangle with circular corners. Oversized radii are clamped during rendering.</summary>
    public OdgShape AddRoundedRectangle(OdfRect bounds, OdfLength radius, string? name = null) {
        OdgShape.ValidateRadius(radius);
        OdgShape shape = AddRectangle(bounds, name); shape.CornerRadius = radius; return shape;
    }
    /// <summary>Adds a path whose SVG data uses the supplied local view box and is scaled to the page bounds.</summary>
    public OdgShape AddPath(OdfRect bounds, OdfViewBox viewBox, string pathData, string? name = null) {
        OdgShape.ParsePath(pathData);
        return AddLocalGeometry("path", bounds, viewBox, OdfNamespaces.Svg + "d", pathData, name);
    }
    /// <summary>Adds a path using shared Drawing commands in the supplied local view box.</summary>
    public OdgShape AddPath(OdfRect bounds, OdfViewBox viewBox, IEnumerable<OfficePathCommand> commands, string? name = null) =>
        AddPath(bounds, viewBox, OdgShape.FormatCommands(commands), name);
    /// <summary>Adds a closed polygon using integer coordinates in the supplied local view box.</summary>
    public OdgShape AddPolygon(OdfRect bounds, OdfViewBox viewBox, IEnumerable<OfficePoint> points, string? name = null) =>
        AddLocalGeometry("polygon", bounds, viewBox, OdfNamespaces.Draw + "points", OdgShape.FormatPoints(points, true), name);
    /// <summary>Adds an open polyline using integer coordinates in the supplied local view box; its initial fill is none.</summary>
    public OdgShape AddPolyline(OdfRect bounds, OdfViewBox viewBox, IEnumerable<OfficePoint> points, string? name = null) {
        OdgShape shape = AddLocalGeometry("polyline", bounds, viewBox, OdfNamespaces.Draw + "points", OdgShape.FormatPoints(points, false), name);
        shape.FillColor = null; return shape;
    }
    private OdgShape AddLocalGeometry(string kind, OdfRect bounds, OdfViewBox viewBox, XName dataName, string data, string? name) {
        viewBox.Validate();
        if (!bounds.X.TryToPoints(out _) || !bounds.Y.TryToPoints(out _) || !bounds.Width.TryToPoints(out double width) || !bounds.Height.TryToPoints(out double height) || width < 0 || height < 0)
            throw new ArgumentException("Geometry requires absolute finite bounds with nonnegative dimensions.", nameof(bounds));
        var element = NewElement(kind, name); OdfShape.ApplyBounds(element, bounds);
        element.SetAttributeValue(OdfNamespaces.Svg + "viewBox", viewBox); element.SetAttributeValue(dataName, data);
        OdgShape shape = Append(element); shape.FillColor = OdfColor.Parse("FFFFFF"); shape.StrokeColor = OdfColor.Parse("000000"); shape.StrokeWidth = OdfLength.Points(1); return shape;
    }
}
