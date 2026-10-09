using OfficeIMO.Drawing;

namespace OfficeIMO.OpenDocument;

public sealed partial class OdgShape {
    internal const int MaximumGeometryItems = 20000;
    private const int MaximumGeometryText = 1024 * 1024;
    internal bool HasLocalGeometry => ElementName is "path" or "polygon" or "polyline";

    /// <summary>Local coordinate canvas for a path, polygon, or polyline. Changing bounds scales this canvas.</summary>
    public OdfViewBox ViewBox {
        get { RequireLocalGeometry(); return OdfViewBox.Parse((string?)Element.Attribute(OdfNamespaces.Svg + "viewBox")); }
        set { RequireLocalGeometry(); value.Validate(); Element.SetAttributeValue(OdfNamespaces.Svg + "viewBox", value); Dirty(); }
    }
    /// <summary>Native SVG path data. Assignment validates the bounded shared path grammar and retains the supplied notation.</summary>
    public string PathData {
        get { RequirePath(); return (string?)Element.Attribute(OdfNamespaces.Svg + "d") ?? string.Empty; }
        set { RequirePath(); ParsePath(value); Element.SetAttributeValue(OdfNamespaces.Svg + "d", value); Dirty(); }
    }
    /// <summary>Parsed path commands in view-box coordinates; SVG arcs and shorthand are normalized by the shared Drawing parser.</summary>
    public IReadOnlyList<OfficePathCommand> PathCommands => ParsePath(PathData);
    /// <summary>Replaces native path data using bounded shared Drawing commands.</summary>
    public void SetPathCommands(IEnumerable<OfficePathCommand> commands) {
        RequirePath(); PathData = FormatCommands(commands);
    }
    /// <summary>Polygon or polyline vertices in integer view-box coordinates. The returned list is detached.</summary>
    public IReadOnlyList<OfficePoint> Points {
        get {
            RequirePoints();
            string value = (string?)Element.Attribute(OdfNamespaces.Draw + "points") ?? string.Empty;
            if (value.Length > MaximumGeometryText) throw new InvalidDataException("Polygon data exceeds 1 MiB.");
            var points = new List<OfficePoint>();
            foreach (string pair in value.Split(new[] { ' ', '\t', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries)) {
                if (points.Count == MaximumGeometryItems) throw new InvalidDataException("Polygon data exceeds 20,000 vertices.");
                string[] xy = pair.Split(',');
                if (xy.Length != 2) throw new FormatException("An ODF vertex requires two comma-separated integers.");
                points.Add(new OfficePoint(int.Parse(xy[0], NumberStyles.AllowLeadingSign, CultureInfo.InvariantCulture), int.Parse(xy[1], NumberStyles.AllowLeadingSign, CultureInfo.InvariantCulture)));
            }
            ValidatePointCount(points.Count, ElementName == "polygon");
            return points.AsReadOnly();
        }
    }
    /// <summary>Replaces polygon/polyline vertices. Coordinates must be signed 32-bit integers; at most 20,000 vertices are accepted.</summary>
    public void SetPoints(IEnumerable<OfficePoint> points) {
        RequirePoints(); string value = FormatPoints(points, ElementName == "polygon");
        Element.SetAttributeValue(OdfNamespaces.Draw + "points", value); Dirty();
    }
    /// <summary>Circular rectangle corner radius. Setting this clears the alternate elliptical radii.</summary>
    public OdfLength? CornerRadius {
        get { RequireRectangle(); return ReadRadius(OdfNamespaces.Draw + "corner-radius"); }
        set { SetRadius(OdfNamespaces.Draw + "corner-radius", value); }
    }
    /// <summary>Horizontal rectangle corner radius. When only one elliptical radius is present it supplies both axes.</summary>
    public OdfLength? CornerRadiusX {
        get { RequireRectangle(); return ReadRadius(OdfNamespaces.Svg + "rx"); }
        set { SetRadius(OdfNamespaces.Svg + "rx", value); }
    }
    /// <summary>Vertical rectangle corner radius. Setting this clears the circular radius.</summary>
    public OdfLength? CornerRadiusY {
        get { RequireRectangle(); return ReadRadius(OdfNamespaces.Svg + "ry"); }
        set { SetRadius(OdfNamespaces.Svg + "ry", value); }
    }
    private void RequireLocalGeometry() { if (!HasLocalGeometry) throw new InvalidOperationException("Only paths, polygons and polylines have editable local geometry."); }
    private void RequirePath() { if (ElementName != "path") throw new InvalidOperationException("Only paths have path commands."); }
    private void RequirePoints() { if (ElementName is not ("polygon" or "polyline")) throw new InvalidOperationException("Only polygons and polylines have vertices."); }
    private void RequireRectangle() { if (ElementName != "rect") throw new InvalidOperationException("Corner radii are supported on rectangles."); }
    private OdfLength? ReadRadius(XName name) => Element.Attribute(name) is XAttribute attribute ? OdfLength.Parse(attribute.Value) : null;
    private void SetRadius(XName name, OdfLength? value) {
        RequireRectangle(); if (value.HasValue) ValidateRadius(value.Value);
        if (value.HasValue) {
            if (name == OdfNamespaces.Draw + "corner-radius") { Element.Attribute(OdfNamespaces.Svg + "rx")?.Remove(); Element.Attribute(OdfNamespaces.Svg + "ry")?.Remove(); }
            else Element.Attribute(OdfNamespaces.Draw + "corner-radius")?.Remove();
        }
        Element.SetAttributeValue(name, value?.ToString()); Dirty();
    }
    internal static void ValidateRadius(OdfLength radius) {
        if (!radius.TryToPoints(out double points) || points < 0) throw new ArgumentOutOfRangeException(nameof(radius), "Corner radii must be finite nonnegative absolute lengths.");
    }
    internal static IReadOnlyList<OfficePathCommand> ParsePath(string? value) {
        if (value == null) throw new ArgumentNullException(nameof(value));
        if (value.Length > MaximumGeometryText) throw new InvalidDataException("Path data exceeds 1 MiB.");
        if (!OfficeSvgPathDataParser.TryParse(value, MaximumGeometryItems, out var commands, out bool exceeded))
            throw new InvalidDataException(exceeded ? "Path data exceeds 20,000 commands." : "Invalid or unsupported SVG path data.");
        OfficeShape.Path(1, 1, commands);
        return commands;
    }
    internal static string FormatCommands(IEnumerable<OfficePathCommand> commands) {
        if (commands == null) throw new ArgumentNullException(nameof(commands));
        var snapshot = commands.Take(MaximumGeometryItems + 1).ToList();
        if (snapshot.Count > MaximumGeometryItems) throw new InvalidDataException("Path data exceeds 20,000 commands.");
        // The shared descriptor checks command order, finite coordinates and drawable content.
        OfficeShape.Path(1, 1, snapshot);
        return OfficeSvgFormatting.FormatEditablePathData(snapshot);
    }
    internal static string FormatPoints(IEnumerable<OfficePoint> points, bool closed) {
        if (points == null) throw new ArgumentNullException(nameof(points));
        var snapshot = points.Take(MaximumGeometryItems + 1).ToList();
        ValidatePointCount(snapshot.Count, closed);
        foreach (OfficePoint point in snapshot) {
            foreach (double coordinate in new[] { point.X, point.Y })
                if (double.IsNaN(coordinate) || coordinate < int.MinValue || coordinate > int.MaxValue || coordinate != Math.Truncate(coordinate))
                    throw new ArgumentException("ODF polygon vertices must be signed 32-bit integers.", nameof(points));
        }
        return string.Join(" ", snapshot.Select(point => ((int)point.X).ToString(CultureInfo.InvariantCulture) + "," + ((int)point.Y).ToString(CultureInfo.InvariantCulture)));
    }
    private static void ValidatePointCount(int count, bool closed) {
        if (count < (closed ? 3 : 2)) throw new InvalidDataException(closed ? "A polygon requires three vertices." : "A polyline requires two vertices.");
        if (count > MaximumGeometryItems) throw new InvalidDataException("Polygon data exceeds 20,000 vertices.");
    }
}
