using OfficeIMO.Drawing;

namespace OfficeIMO.OpenDocument;

public sealed partial class OdgShape {
    /// <summary>
    /// Returns a detached route in connector-local points, before the connector and parent transforms.
    /// Saved routes must match the resolved endpoints; non-straight connectors without a saved route require native routing.
    /// </summary>
    public IReadOnlyList<OfficePathCommand> ConnectorRouteCommands {
        get {
            RequireConnector();
            OfficePoint start = new(X1.ToPoints(), Y1.ToPoints()), end = new(X2.ToPoints(), Y2.ToPoints());
            if (!Finite(start.X) || !Finite(start.Y) || !Finite(end.X) || !Finite(end.Y)) throw new InvalidDataException("Connector endpoints must be finite.");
            string? data = (string?)Element.Attribute(OdfNamespaces.Svg + "d");
            if (data == null) {
                if (ConnectorKind != OdgConnectorKind.Line) throw new NotSupportedException("Connector has no saved route; native routing is required.");
                return new[] { OfficePathCommand.MoveTo(start), OfficePathCommand.LineTo(end) };
            }
            IReadOnlyList<OfficePathCommand> native = ParsePath(data);
            ValidateConnectorPath(native);
            OdfViewBox box = OdfViewBox.Parse((string?)Element.Attribute(OdfNamespaces.Svg + "viewBox"));
            // LibreOffice/OpenOffice connector caches use page coordinates in 1/100 mm,
            // independently of the view box's advertised width and height.
            var points = native.Select(command => command.Scale(72D / 2540, 72D / 2540)).ToArray();
            if (RouteMatches(points, start, end)) return Array.AsReadOnly(points);
            // Standards-based producers may instead use the declared local coordinate canvas.
            if (TryReadSavedEndpoint(true, out OfficePoint savedStart) && TryReadSavedEndpoint(false, out OfficePoint savedEnd)) {
                double x = Math.Min(savedStart.X, savedEnd.X), y = Math.Min(savedStart.Y, savedEnd.Y);
                var mapped = native.Select(command => command.Translate(box.X, box.Y)
                    .Scale(Math.Abs(savedEnd.X - savedStart.X) / box.Width, Math.Abs(savedEnd.Y - savedStart.Y) / box.Height)
                    .Translate(-x, -y)).ToArray();
                if (RouteMatches(mapped, start, end)) return Array.AsReadOnly(mapped);
            }
            throw new NotSupportedException("Saved connector route does not match its current endpoints. Replace the route or recalculate it in a native drawing application.");
        }
    }

    /// <summary>
    /// Replaces the saved open route with commands in connector-local points. Free endpoints follow the route;
    /// attached endpoints within 0.1 point of their glue positions are snapped to those positions before native-grid rounding.
    /// Native applications can subsequently reroute it.
    /// </summary>
    public void SetConnectorRoute(IEnumerable<OfficePathCommand> commands) => SetConnectorRouteCore(commands, null);

    private void SetConnectorRouteCore(IEnumerable<OfficePathCommand> commands, OdgConnectorKind? routingKind) {
        RequireConnector();
        if (commands == null) throw new ArgumentNullException(nameof(commands));
        var path = commands.Take(MaximumGeometryItems + 1).ToArray();
        _ = FormatCommands(path); ValidateConnectorPath(path);
        OfficePoint start = path[0].Point, end = path[path.Length - 1].Point;
        if ((StartShapeId != null && !Near(start, new OfficePoint(X1.ToPoints(), Y1.ToPoints()))) ||
            (EndShapeId != null && !Near(end, new OfficePoint(X2.ToPoints(), Y2.ToPoints()))))
            throw new ArgumentException("Route endpoints must match attached glue positions.", nameof(commands));
        if ((routingKind ?? ConnectorKind) == OdgConnectorKind.Line && (path.Length != 2 || path[1].Kind != OfficePathCommandKind.LineTo))
            throw new ArgumentException("A line connector requires one straight segment. Select another routing kind before assigning bends or curves.", nameof(commands));
        if (StartShapeId != null) path[0] = OfficePathCommand.MoveTo(X1.ToPoints(), Y1.ToPoints());
        if (EndShapeId != null) path[path.Length - 1] = ReplaceConnectorEndpoint(path[path.Length - 1], new OfficePoint(X2.ToPoints(), Y2.ToPoints()));
        var attributes = EncodeConnectorRoute(path);
        if (routingKind.HasValue) Element.SetAttributeValue(OdfNamespaces.Draw + "type", routingKind.Value.ToString().ToLowerInvariant());
        foreach (XAttribute attribute in attributes) Element.SetAttributeValue(attribute.Name, attribute.Value);
        Element.Attribute(OdfNamespaces.Draw + "line-skew")?.Remove(); Dirty();
    }

    private static OfficePathCommand ReplaceConnectorEndpoint(OfficePathCommand command, OfficePoint point) => command.Kind switch {
        OfficePathCommandKind.LineTo => OfficePathCommand.LineTo(point),
        OfficePathCommandKind.QuadraticBezierTo => OfficePathCommand.QuadraticBezierTo(command.ControlPoint1, point),
        OfficePathCommandKind.CubicBezierTo => OfficePathCommand.CubicBezierTo(command.ControlPoint1, command.ControlPoint2, point),
        _ => throw new NotSupportedException("A connector route must end with a drawing command.")
    };

    private static XAttribute[] EncodeConnectorRoute(IReadOnlyList<OfficePathCommand> path) {
        var native = path.Select(command => MapConnectorCommand(command, NativeConnectorPoint)).ToArray();
        string data = FormatCommands(native);
        _ = ParsePath(data);
        var positions = ConnectorPathPoints(native).ToArray();
        if (positions.Any(point => point.X < int.MinValue || point.X > int.MaxValue || point.Y < int.MinValue || point.Y > int.MaxValue))
            throw new ArgumentOutOfRangeException(nameof(path), "Route coordinates exceed signed 32-bit native coordinates.");
        double width = Math.Ceiling(positions.Max(point => point.X) - positions.Min(point => point.X)) + 1;
        double height = Math.Ceiling(positions.Max(point => point.Y) - positions.Min(point => point.Y)) + 1;
        if (width > int.MaxValue || height > int.MaxValue) throw new ArgumentOutOfRangeException(nameof(path), "Route canvas exceeds signed 32-bit native coordinates.");
        OfficePoint start = native[0].Point, end = native[native.Length - 1].Point;
        return new[] {
            new XAttribute(OdfNamespaces.Svg + "viewBox", new OdfViewBox(0, 0, (int)width, (int)height)),
            new XAttribute(OdfNamespaces.Svg + "x1", OdfLength.Centimeters(start.X / 1000)),
            new XAttribute(OdfNamespaces.Svg + "y1", OdfLength.Centimeters(start.Y / 1000)),
            new XAttribute(OdfNamespaces.Svg + "x2", OdfLength.Centimeters(end.X / 1000)),
            new XAttribute(OdfNamespaces.Svg + "y2", OdfLength.Centimeters(end.Y / 1000)),
            new XAttribute(OdfNamespaces.Svg + "d", data)
        };
    }

    private static OfficePoint NativeConnectorPoint(OfficePoint point) => new(
        Math.Round(point.X * 2540 / 72, MidpointRounding.AwayFromZero), Math.Round(point.Y * 2540 / 72, MidpointRounding.AwayFromZero));

    private static OfficePathCommand MapConnectorCommand(OfficePathCommand command, Func<OfficePoint, OfficePoint> map) => command.Kind switch {
        OfficePathCommandKind.MoveTo => OfficePathCommand.MoveTo(map(command.Point)),
        OfficePathCommandKind.LineTo => OfficePathCommand.LineTo(map(command.Point)),
        OfficePathCommandKind.QuadraticBezierTo => OfficePathCommand.QuadraticBezierTo(map(command.ControlPoint1), map(command.Point)),
        OfficePathCommandKind.CubicBezierTo => OfficePathCommand.CubicBezierTo(map(command.ControlPoint1), map(command.ControlPoint2), map(command.Point)),
        _ => throw new NotSupportedException("A connector route must be one open path.")
    };

    internal static IEnumerable<OfficePoint> ConnectorPathPoints(IEnumerable<OfficePathCommand> commands) {
        foreach (OfficePathCommand command in commands) {
            yield return command.Point;
            if (command.Kind is OfficePathCommandKind.QuadraticBezierTo or OfficePathCommandKind.CubicBezierTo) yield return command.ControlPoint1;
            if (command.Kind == OfficePathCommandKind.CubicBezierTo) yield return command.ControlPoint2;
        }
    }
    private static void ValidateConnectorPath(IReadOnlyList<OfficePathCommand> path) {
        if (path.Count < 2 || path[0].Kind != OfficePathCommandKind.MoveTo ||
            path.Skip(1).Any(command => command.Kind is OfficePathCommandKind.MoveTo or OfficePathCommandKind.Close))
            throw new NotSupportedException("A connector route must be one open path with at least one drawing command.");
    }
    private static bool RouteMatches(IReadOnlyList<OfficePathCommand> path, OfficePoint start, OfficePoint end) =>
        Near(path[0].Point, start) && Near(path[path.Count - 1].Point, end);
}
