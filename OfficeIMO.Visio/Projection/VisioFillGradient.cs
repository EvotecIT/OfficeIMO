using System.Globalization;
using System.Xml.Linq;
using OfficeIMO.Drawing;

namespace OfficeIMO.Visio;

/// <summary>Cached native gradient fill of a shape, resolved through its master without evaluating formulas.</summary>
internal sealed class VisioFillGradient {
    private VisioFillGradient(int direction, double angle, IReadOnlyList<OfficeGradientStop> stops) {
        Direction = direction; Angle = angle; Stops = stops;
    }

    /// <summary>FillGradientDir: 0 linear, 1-7 radial, 8-12 rectangular, 13 path.</summary>
    internal int Direction { get; }

    /// <summary>FillGradientAngle in radians, clockwise from left-to-right in page orientation.</summary>
    internal double Angle { get; }

    internal IReadOnlyList<OfficeGradientStop> Stops { get; }

    internal bool IsLinear => Direction == 0;

    /// <summary>Reads an enabled gradient. Returns false when the shape has none; <paramref name="unreadable"/> is set for an enabled gradient whose stops are not cached literals.</summary>
    internal static bool TryResolve(VisioShape shape, out VisioFillGradient? gradient, out bool unreadable) {
        gradient = null; unreadable = false;
        if (Cell(shape, "FillGradientEnabled") is not string enabled) return false;
        // Theme-driven values are outside the cached literal profile, like other theme formatting.
        if (!VisioShapeGeometry.TryParseLiteralWithoutShape(enabled, out double flag)) return false;
        if (flag == 0) return false;
        int direction = VisioShapeGeometry.TryParseLiteralWithoutShape(Cell(shape, "FillGradientDir"), out double rawDirection)
            && rawDirection >= 0 && rawDirection <= 13 ? (int)rawDirection : 0;
        double angle = VisioShapeGeometry.TryParseLiteralWithoutShape(Cell(shape, "FillGradientAngle"), out double rawAngle)
            && !double.IsNaN(rawAngle) && !double.IsInfinity(rawAngle) ? rawAngle : Math.PI / 2D;
        XElement? section = Section(shape, "FillGradient");
        var stops = new List<OfficeGradientStop>();
        if (section != null) {
            XNamespace ns = section.Name.Namespace;
            double previous = 0D;
            foreach (XElement row in section.Elements(ns + "Row")) {
                if ((string?)row.Attribute("Del") == "1") continue;
                string? color = RowCell(row, "GradientStopColor");
                if (color == null || !OfficeColor.TryParse(color, out OfficeColor parsed)) { unreadable = true; return false; }
                double position = VisioShapeGeometry.TryParseLiteralWithoutShape(RowCell(row, "GradientStopPosition"), out double rawPosition)
                    ? rawPosition : stops.Count == 0 ? 0D : 1D;
                position = Math.Max(previous, Math.Min(1D, Math.Max(0D, position)));
                if (VisioShapeGeometry.TryParseLiteralWithoutShape(RowCell(row, "GradientStopColorTrans"), out double transparency)
                    && transparency > 0 && transparency <= 1)
                    parsed = OfficeColor.FromRgba(parsed.R, parsed.G, parsed.B, (byte)Math.Round(parsed.A * (1D - transparency)));
                stops.Add(new OfficeGradientStop(position, parsed));
                previous = position;
            }
        }
        if (stops.Count < 2) { unreadable = true; return false; }
        // Visio pads before the first and after the last stop; the shared model spans the full 0..1 range.
        if (stops[0].Offset > 0D) stops.Insert(0, new OfficeGradientStop(0D, stops[0].Color));
        if (stops[stops.Count - 1].Offset < 1D) stops.Add(new OfficeGradientStop(1D, stops[stops.Count - 1].Color));
        gradient = new VisioFillGradient(direction, angle, stops);
        return true;
    }

    /// <summary>
    /// Places the gradient in a renderer's coordinates. <paramref name="map"/> takes local shape inches.
    /// A linear gradient yields its start and end points; FillGradientAngle follows the Office convention,
    /// clockwise from left-to-right as displayed, and the line spans the frame's projected extent. Other
    /// directions yield a frame whose unit square maps the gradient circle (centre 0.5, radius 0.5) onto an
    /// ellipse reaching the corners of the shape's local frame.
    /// </summary>
    internal bool TryPlace(VisioShape shape, Func<double, double, OfficePoint> map,
        out OfficePoint start, out OfficePoint end, out OfficeTransform radialFrame) {
        double width = shape.WidthInches, height = shape.HeightInches;
        start = end = default; radialFrame = OfficeTransform.Identity;
        if (IsLinear) {
            double dx = Math.Cos(Angle), dy = -Math.Sin(Angle);
            double half = (Math.Abs(width * dx) + Math.Abs(height * dy)) / 2D;
            start = map(width / 2D - dx * half, height / 2D - dy * half);
            end = map(width / 2D + dx * half, height / 2D + dy * half);
            return start.X != end.X || start.Y != end.Y;
        }
        double reach = Math.Sqrt(2D);
        OfficePoint origin = map(width / 2D - reach * width / 2D, height / 2D + reach * height / 2D);
        OfficePoint unitX = map(width / 2D + reach * width / 2D, height / 2D + reach * height / 2D);
        OfficePoint unitY = map(width / 2D - reach * width / 2D, height / 2D - reach * height / 2D);
        radialFrame = new OfficeTransform(unitX.X - origin.X, unitX.Y - origin.Y, unitY.X - origin.X, unitY.Y - origin.Y, origin.X, origin.Y);
        return radialFrame.TryInvert(out _);
    }

    /// <summary>Creates a per-pixel color function for raster fills in the same coordinates as <paramref name="map"/>.</summary>
    internal Func<double, double, OfficeColor>? CreatePaint(VisioShape shape, Func<double, double, OfficePoint> map) {
        if (!TryPlace(shape, map, out OfficePoint start, out OfficePoint end, out OfficeTransform frame)) return null;
        IReadOnlyList<OfficeGradientStop> stops = Stops;
        if (IsLinear) {
            double dx = end.X - start.X, dy = end.Y - start.Y, length = dx * dx + dy * dy;
            return (x, y) => OfficeRasterCanvas.InterpolateGradientStops(stops, Clamp(((x - start.X) * dx + (y - start.Y) * dy) / length));
        }
        OfficeTransform inverse = frame.Invert();
        return (x, y) => {
            OfficePoint unit = inverse.TransformPoint(new OfficePoint(x, y));
            double ux = unit.X - 0.5D, uy = unit.Y - 0.5D;
            return OfficeRasterCanvas.InterpolateGradientStops(stops, Clamp(Math.Sqrt(ux * ux + uy * uy) / 0.5D));
        };
    }

    private static double Clamp(double value) => double.IsNaN(value) ? 0D : Math.Max(0D, Math.Min(1D, value));

    private static IEnumerable<VisioShape> Chain(VisioShape shape) {
        VisioShape? current = shape;
        for (int depth = 0; current != null && depth < 8; depth++) {
            yield return current;
            VisioShape? next = current.MasterShape ?? current.Master?.Shape;
            current = ReferenceEquals(next, current) ? null : next;
        }
    }

    private static string? Cell(VisioShape shape, string name) {
        foreach (VisioShape source in Chain(shape))
            foreach (XElement cell in source.PreservedCellElements)
                if ((string?)cell.Attribute("N") == name) return (string?)cell.Attribute("V");
        return null;
    }

    private static XElement? Section(VisioShape shape, string name) {
        foreach (VisioShape source in Chain(shape))
            foreach (XElement section in source.PreservedNonGeometrySections)
                if ((string?)section.Attribute("N") == name) return section;
        return null;
    }

    private static string? RowCell(XElement row, string name) =>
        row.Elements(row.Name.Namespace + "Cell").FirstOrDefault(cell => (string?)cell.Attribute("N") == name)?.Attribute("V")?.Value;

    internal static string Describe(int direction) => direction switch {
        0 => "linear",
        <= 7 => "radial",
        <= 12 => "rectangular",
        _ => "path"
    } + " (" + direction.ToString(CultureInfo.InvariantCulture) + ")";
}
