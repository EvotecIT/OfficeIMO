using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using OfficeIMO.Drawing;

namespace OfficeIMO.Visio;

public partial class VisioConnector {
    internal VisioConnectorNativeGeometry? NativeGeometry { get; set; }
}

/// <summary>Retains a native local coordinate system and fits it to edited page endpoints.</summary>
internal sealed class VisioConnectorNativeGeometry {
    private readonly VisioShape _source;
    private readonly OfficePoint _start, _end;
    private readonly ConnectorKind _kind;
    private readonly bool _flipX, _flipY;
    private List<(double X, double Y)> _waypoints = new();

    internal VisioConnectorNativeGeometry(VisioShape source, OfficePoint start, OfficePoint end, ConnectorKind kind, bool flipX, bool flipY) {
        _source = source; _start = start; _end = end; _kind = kind; _flipX = flipX; _flipY = flipY;
    }

    internal bool FlipX => _flipX;
    internal bool FlipY => _flipY;

    internal void CaptureRoute(VisioConnector connector) => _waypoints = connector.Waypoints.Select(p => (p.X, p.Y)).ToList();
    internal VisioConnectorNativeGeometry CopyFor(VisioConnector connector) {
        var copy = new VisioConnectorNativeGeometry(_source, _start, _end, _kind, _flipX, _flipY);
        copy.CaptureRoute(connector);
        return copy;
    }
    internal bool AppliesTo(VisioConnector connector) => connector.PreservedGeometrySections.Count > 0 && connector.Kind == _kind &&
        _waypoints.SequenceEqual(connector.Waypoints.Select(p => (p.X, p.Y))) &&
        (_waypoints.Count == 0 || (SamePoint(connector.StartPoint, _start) && SamePoint(connector.EndPoint, _end)));

    private static bool SamePoint(OfficePoint a, OfficePoint b) => Math.Abs(a.X - b.X) < 1e-9 && Math.Abs(a.Y - b.Y) < 1e-9;
    internal (double X, double Y) SourceToPage(double x, double y) => ToPage(_source, x, y);

    /// <summary>The same similarity transform used by the preserved curve and its cached text frame.</summary>
    internal (double Scale, double Angle) GetTransform(VisioConnector connector) {
        OfficePoint start = connector.StartPoint, end = connector.EndPoint;
        double dx = _end.X - _start.X, dy = _end.Y - _start.Y;
        double length = Math.Sqrt(dx * dx + dy * dy);
        double newDx = end.X - start.X, newDy = end.Y - start.Y;
        double newLength = Math.Sqrt(newDx * newDx + newDy * newDy);
        if (length < 1e-12 && newLength >= 1e-12)
            throw new NotSupportedException("A closed native connector cannot be stretched by separating its endpoints. Replace its route first.");
        double scale = length < 1e-12 ? 1 : newLength / length;
        double angle = length < 1e-12 ? 0 : Math.Atan2(newDy, newDx) - Math.Atan2(dy, dx);
        return (scale, angle);
    }

    internal OfficePoint TransformSourcePoint(VisioConnector connector, OfficePoint point) {
        (double scale, double angle) = GetTransform(connector);
        double x = (point.X - _start.X) * scale, y = (point.Y - _start.Y) * scale;
        OfficePoint start = connector.StartPoint;
        return new OfficePoint(start.X + x * Math.Cos(angle) - y * Math.Sin(angle), start.Y + x * Math.Sin(angle) + y * Math.Cos(angle));
    }

    internal VisioShape CreateShape(VisioConnector connector, bool requireCompleteRows = false) {
        OfficePoint start = connector.StartPoint, end = connector.EndPoint;
        (double scale, double angle) = GetTransform(connector);
        double cosine = Math.Cos(angle), sine = Math.Sin(angle);
        double px = (_source.PinX - _start.X) * scale, py = (_source.PinY - _start.Y) * scale;
        VisioShape shape = new("connector-geometry") {
            PinX = start.X + px * cosine - py * sine, PinY = start.Y + px * sine + py * cosine,
            Width = _source.Width * scale, Height = _source.Height * scale,
            LocPinX = _source.LocPinX * scale, LocPinY = _source.LocPinY * scale, Angle = _source.Angle + angle
        };
        foreach (XElement section in connector.PreservedGeometrySections) shape.PreservedGeometrySections.Add(new XElement(section));
        VisioGeometryScaling.Scale(shape, scale, scale, requireCompleteRows: requireCompleteRows);
        return shape;
    }

    internal (double X, double Y) ToPage(VisioShape shape, double x, double y) =>
        shape.GetAbsolutePoint(_flipX ? shape.Width - x : x, _flipY ? shape.Height - y : y);

    internal bool TryGetPoints(VisioConnector connector, out List<(double X, double Y)> points) {
        points = new();
        if (!AppliesTo(connector)) return false;
        VisioShape shape = CreateShape(connector);
        if (!VisioShapeGeometry.TryGetPreservedClosedPaths(shape, out var paths, includeHidden: true) || paths.Count != 1) {
            // With multiple structural paths, only a unique visible outline identifies the route.
            if (!VisioShapeGeometry.TryGetPreservedClosedPaths(shape, out paths)) return false;
            paths.RemoveAll(path => path.NoLine);
            if (paths.Count != 1) return false;
        }
        points = paths[0].Points.Select(p => ToPage(shape, p.X, p.Y)).ToList();
        return points.Count >= 2;
    }

    internal bool TryHasVisibleLine(VisioConnector connector, out bool visible) {
        visible = true;
        if (!AppliesTo(connector)) return false;
        VisioShape shape = CreateShape(connector);
        visible = VisioShapeGeometry.TryGetPreservedClosedPaths(shape, out var paths)
            ? paths.Any(path => !path.NoLine)
            : VisioShapeGeometry.HasVisibleSectionLine(shape);
        return true;
    }
}

/// <summary>Shared path projection for connector drawing, labels, inspection and intersection tests.</summary>
internal static class VisioConnectorGeometry {
    internal static bool IsTransformCell(string? name) => name == "PinX" || name == "PinY" || name == "Width" || name == "Height" ||
        name == "LocPinX" || name == "LocPinY" || name == "Angle" || name == "FlipX" || name == "FlipY";

    internal static List<(double X, double Y)> GetPoints(VisioConnector connector) {
        if (connector.NativeGeometry?.TryGetPoints(connector, out var points) == true) return points;
        OfficePoint start = connector.StartPoint, end = connector.EndPoint;
        return OfficeGeometry.BuildConnectorPolyline((start.X, start.Y), (end.X, end.Y),
            connector.Waypoints.Select(p => (p.X, p.Y)).ToList(), connector.Kind == ConnectorKind.RightAngle);
    }

    /// <summary>Resolves outline visibility separately from routes used by labels, routing and inspection.</summary>
    internal static bool HasVisibleLine(VisioConnector connector) =>
        connector.LinePattern != 0 && connector.LineWeight > 0 && connector.LineColor.A > 0 &&
        (connector.NativeGeometry == null || !connector.NativeGeometry.TryHasVisibleLine(connector, out bool visible) || visible);

    /// <summary>Transforms the authored text box center without moving it when a renderer expands its fitting box.</summary>
    internal static (double X, double Y) GetLabelCenter(VisioConnector connector, double pinX, double pinY, double width, double height) {
        VisioConnectorLabelPlacement? placement = VisioConnectorLabelFrame.ResolvePlacement(connector);
        double authoredWidth = connector.TextStyle?.TextWidth ?? placement?.Width ?? width;
        double authoredHeight = connector.TextStyle?.TextHeight ?? placement?.Height ?? height;
        double dx = authoredWidth / 2 - (placement?.LocPinX ?? connector.TextStyle?.TextLocPinX ?? authoredWidth / 2);
        double dy = authoredHeight / 2 - (placement?.LocPinY ?? connector.TextStyle?.TextLocPinY ?? authoredHeight / 2);
        double angle = VisioConnectorLabelFrame.ResolveAngle(connector);
        return (pinX + dx * Math.Cos(angle) - dy * Math.Sin(angle), pinY + dx * Math.Sin(angle) + dy * Math.Cos(angle));
    }

    internal static VisioShape CreateShape(VisioConnector connector) {
        if (connector.NativeGeometry?.AppliesTo(connector) == true) return connector.NativeGeometry.CreateShape(connector);
        OfficePoint start = connector.StartPoint, end = connector.EndPoint;
        double dx = end.X - start.X, dy = end.Y - start.Y, width = Math.Sqrt(dx * dx + dy * dy);
        VisioShape shape = new("connector-geometry") {
            PinX = (start.X + end.X) / 2, PinY = (start.Y + end.Y) / 2, Width = width, Height = 0,
            LocPinX = width / 2, LocPinY = 0, Angle = Math.Atan2(dy, dx)
        };
        XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
        XElement section = VisioGeometrySection.CreateGenerated(ns, noFill: true);
        int rowIndex = 0;
        foreach (var point in GetPoints(connector)) {
            OfficePoint local = VisioConnectorEndpoints.ToLocalPoint(shape, new OfficePoint(point.X, point.Y));
            section.Add(new XElement(ns + "Row", new XAttribute("T", rowIndex == 0 ? "MoveTo" : "LineTo"), new XAttribute("IX", ++rowIndex),
                new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", local.X)),
                new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", local.Y))));
        }
        shape.PreservedGeometrySections.Add(section);
        return shape;
    }
}
