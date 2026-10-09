using OfficeIMO.Drawing;

namespace OfficeIMO.OpenDocument;

public sealed partial class OdgShape {
    internal OfficeShape CreateDrawingGeometry(double width, double height) {
        if (double.IsNaN(width) || double.IsInfinity(width) || double.IsNaN(height) || double.IsInfinity(height) || width < 0 || height < 0)
            throw new InvalidDataException("Shape dimensions must be finite and nonnegative.");
        if (ElementName == "custom-shape") return OdfEnhancedGeometry.Project(Element, width, height);
        if (HasLocalGeometry) {
            OdfViewBox box = ViewBox;
            IEnumerable<OfficePathCommand> commands;
            if (ElementName == "path") commands = PathCommands;
            else {
                var vertices = Points;
                var path = vertices.Select((point, index) => index == 0 ? OfficePathCommand.MoveTo(point) : OfficePathCommand.LineTo(point)).ToList();
                if (ElementName == "polygon") path.Add(OfficePathCommand.Close());
                commands = path;
            }
            return OfficeShape.Path(Math.Max(width, 0.001), Math.Max(height, 0.001),
                OdfViewBoxProjection.Project(commands, box, width, height));
        }
        if (ElementName is "ellipse" or "circle") {
            if (((string?)Element.Attribute(OdfNamespaces.Draw + "kind")) is string kind && kind != "full")
                throw new NotSupportedException("Ellipse and circle arc/sector geometry is outside the full-ellipse projection profile.");
            if (Element.Attribute(OdfNamespaces.Svg + "width") == null || Element.Attribute(OdfNamespaces.Svg + "height") == null)
                throw new NotSupportedException("Ellipse and circle projection requires an explicit bounding width and height.");
            return OfficeShape.Ellipse(width, height);
        }
        if (ElementName != "rect") return OfficeShape.Rectangle(width, height);
        OdfLength? rx = CornerRadiusX, ry = CornerRadiusY;
        foreach (OdfLength? radius in new[] { rx, ry, CornerRadius }) if (radius.HasValue) ValidateRadius(radius.Value);
        double x = (rx ?? ry ?? CornerRadius ?? OdfLength.Points(0)).ToPoints();
        double y = (ry ?? rx ?? CornerRadius ?? OdfLength.Points(0)).ToPoints();
        if (x < 0 || y < 0) throw new InvalidDataException("Rectangle corner radii must be nonnegative.");
        if (!rx.HasValue && !ry.HasValue) x = y = Math.Min(x, Math.Min(width, height) / 2);
        else { x = Math.Min(x, width / 2); y = Math.Min(y, height / 2); }
        if (x == 0 || y == 0) return OfficeShape.Rectangle(width, height);
        if (x == y) return OfficeShape.RoundedRectangle(width, height, x);
        // Shared SVG path parsing owns elliptical arc conversion; keep the declared canvas.
        string F(double value) => value.ToString("R", CultureInfo.InvariantCulture);
        string data = $"M{F(x)} 0H{F(width - x)}A{F(x)} {F(y)} 0 0 1 {F(width)} {F(y)}V{F(height - y)}A{F(x)} {F(y)} 0 0 1 {F(width - x)} {F(height)}H{F(x)}A{F(x)} {F(y)} 0 0 1 0 {F(height - y)}V{F(y)}A{F(x)} {F(y)} 0 0 1 {F(x)} 0Z";
        return OfficeShape.Path(width, height, ParsePath(data));
    }
}
