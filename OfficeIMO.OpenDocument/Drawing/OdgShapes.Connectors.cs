using OfficeIMO.Drawing;

namespace OfficeIMO.OpenDocument;

public sealed partial class OdgShapes {
    /// <summary>Adds a connector between free points in this collection's coordinate space.</summary>
    public OdgShape AddConnector(OfficePoint start, OfficePoint end, OdgConnectorKind kind = OdgConnectorKind.Line, string? name = null) {
        if (!Enum.IsDefined(typeof(OdgConnectorKind), kind)) throw new ArgumentOutOfRangeException(nameof(kind));
        foreach (double coordinate in new[] { start.X, start.Y, end.X, end.Y })
            if (double.IsNaN(coordinate) || double.IsInfinity(coordinate)) throw new ArgumentException("Connector endpoints must be finite.");
        var element = NewElement("connector", name);
        // ODF requires a coordinate canvas even when a straight connector has no cached path.
        element.SetAttributeValue(OdfNamespaces.Svg + "viewBox", "0 0 1 1");
        element.SetAttributeValue(OdfNamespaces.Draw + "type", kind.ToString().ToLowerInvariant());
        element.SetAttributeValue(OdfNamespaces.Svg + "x1", OdfLength.Points(start.X));
        element.SetAttributeValue(OdfNamespaces.Svg + "y1", OdfLength.Points(start.Y));
        element.SetAttributeValue(OdfNamespaces.Svg + "x2", OdfLength.Points(end.X));
        element.SetAttributeValue(OdfNamespaces.Svg + "y2", OdfLength.Points(end.Y));
        var shape = Append(element);
        shape.StrokeColor = OdfColor.Parse("000000"); shape.StrokeWidth = OdfLength.Points(1);
        return shape;
    }
}
