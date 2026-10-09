using System;
using System.Linq;
using System.Xml.Linq;

namespace OfficeIMO.Visio;

internal static partial class VisioGeometryScaling {
    // EllipticalArcTo caches an endpoint, a point on the arc, an axis angle and
    // an axis ratio. Positive diagonal scaling preserves the selected sweep.
    private static void ScaleArc(XElement row, string type, VisioShape shape, VisioShape original,
        (double X, double Y)? start, (double X, double Y)? end, double x, double y) {
        if (x <= 0 || y <= 0 || double.IsInfinity(x) || double.IsInfinity(y) || double.IsNaN(x) || double.IsNaN(y))
            throw new NotSupportedException("Nonuniform arc resizing requires positive finite scale factors.");
        XNamespace ns = row.Name.Namespace;
        if (start == null || end == null)
            throw new NotSupportedException("Nonuniform arc resizing requires readable start and end coordinates.");
        double angle = 0, ratio = 1;
        double controlX, controlY;
        if (type == "ArcTo") {
            if (!VisioShapeGeometry.TryReadCell(row, ns, "A", original, out double bow))
                throw new NotSupportedException("Nonuniform circular arc resizing requires a readable bow.");
            if (bow == 0) {
                MaterializeCell(Cell(row, "X"), end.Value.X * x, shape);
                MaterializeCell(Cell(row, "Y"), end.Value.Y * y, shape);
                MaterializeCell(Cell(row, "A"), 0, shape);
                return;
            }
            double dx = end.Value.X - start.Value.X, dy = end.Value.Y - start.Value.Y;
            double magnitude = Math.Max(Math.Abs(dx), Math.Abs(dy));
            if (magnitude == 0 || double.IsInfinity(magnitude))
                throw new NotSupportedException("Nonuniform circular arc resizing requires a nonzero finite chord.");
            double normalizedLength = Math.Sqrt((dx / magnitude) * (dx / magnitude) + (dy / magnitude) * (dy / magnitude));
            controlX = start.Value.X / 2 + end.Value.X / 2 + (dy / magnitude / normalizedLength) * bow;
            controlY = start.Value.Y / 2 + end.Value.Y / 2 - (dx / magnitude / normalizedLength) * bow;
        } else {
            if (!VisioShapeGeometry.TryReadCell(row, ns, "A", original, out controlX) ||
                !VisioShapeGeometry.TryReadCell(row, ns, "B", original, out controlY))
                throw new NotSupportedException("Nonuniform elliptical arc resizing requires readable control coordinates.");
            if (type == "RelEllipticalArcTo") {
                controlX *= original.Width;
                controlY *= original.Height;
            }
            if (row.Elements(ns + "Cell").Any(c => (string?)c.Attribute("N") == "C") &&
                !VisioShapeGeometry.TryReadCell(row, ns, "C", original, out angle) ||
                row.Elements(ns + "Cell").Any(c => (string?)c.Attribute("N") == "D") &&
                !VisioShapeGeometry.TryReadCell(row, ns, "D", original, out ratio))
                throw new NotSupportedException("Nonuniform elliptical arc resizing requires a readable axis profile.");
        }
        if (!Finite(start.Value.X * x) || !Finite(start.Value.Y * y) || !Finite(end.Value.X * x) ||
            !Finite(end.Value.Y * y) || !Finite(controlX * x) || !Finite(controlY * y))
            throw new NotSupportedException("Nonuniform arc resizing requires finite transformed coordinates.");
        (double Angle, double Ratio) profile = ScaleEllipseAxes(angle, ratio, x, y);
        if (type == "RelEllipticalArcTo") {
            MaterializeRelativeCells(row, original, shape);
        } else {
            MaterializeCell(Cell(row, "X"), end.Value.X * x, shape);
            MaterializeCell(Cell(row, "Y"), end.Value.Y * y, shape);
            MaterializeCell(Cell(row, "A", "IN"), controlX * x, shape, replaceMeaning: type == "ArcTo");
            MaterializeCell(Cell(row, "B", "IN"), controlY * y, shape);
        }
        MaterializeCell(Cell(row, "C", "RAD"), profile.Angle, shape, replaceMeaning: true);
        MaterializeCell(Cell(row, "D"), profile.Ratio, shape, replaceMeaning: true);
        if (type == "ArcTo") row.SetAttributeValue("T", "EllipticalArcTo");
    }

    private static bool Finite(double value) => !double.IsNaN(value) && !double.IsInfinity(value);

    private static (double Angle, double Ratio) ScaleEllipseAxes(double angle, double ratio, double x, double y) {
        if (ratio <= 0 || ratio > 1000)
            throw new NotSupportedException("Nonuniform elliptical arc resizing requires an axis ratio greater than zero and at most 1000.");
        // Principal axes of S R(angle) diag(ratio, 1). Normalize before squaring
        // so large but finite drawing dimensions do not overflow the decomposition.
        double scale = Math.Max(x, y), axis = Math.Max(ratio, 1);
        double sx = x / scale, sy = y / scale, major = ratio / axis, minor = 1 / axis;
        double cos = Math.Cos(angle), sin = Math.Sin(angle);
        double a = sx * sx * (major * major * cos * cos + minor * minor * sin * sin);
        double b = sx * sy * (major * major - minor * minor) * sin * cos;
        double c = sy * sy * (major * major * sin * sin + minor * minor * cos * cos);
        double largest = (a + c + Math.Sqrt((a - c) * (a - c) + 4 * b * b)) / 2;
        double determinant = sx * sy * major * minor;
        double result = largest / determinant;
        if (double.IsNaN(result) || double.IsInfinity(result) || result > 1000 || result < 1 - 1e-12)
            throw new NotSupportedException("The resized elliptical arc exceeds the supported principal-axis ratio of 1000.");
        return (Math.Atan2(2 * b, a - c) / 2, Math.Max(1, result));
    }

    private static XElement Cell(XElement row, string name, string? unit = null) {
        XNamespace ns = row.Name.Namespace;
        XElement? cell = row.Elements(ns + "Cell").FirstOrDefault(c => (string?)c.Attribute("N") == name);
        if (cell != null) return cell;
        cell = new XElement(ns + "Cell", new XAttribute("N", name));
        if (unit != null) cell.SetAttributeValue("U", unit);
        row.Add(cell);
        return cell;
    }
}
