using System;
using System.Globalization;
using System.Linq;
using System.Xml.Linq;

namespace OfficeIMO.Visio;

/// <summary>Materializes resized native geometry without requiring a reader to recalculate ShapeSheet formulas.</summary>
internal static partial class VisioGeometryScaling {
    internal static void Scale(VisioShape shape, double x, double y, VisioShape? originalShape = null,
        bool requireCompleteRows = true) {
        if (Math.Abs(x - 1) < 1e-12 && Math.Abs(y - 1) < 1e-12) return;
        bool uniform = Math.Abs(x - y) < 1e-12;
        VisioShape original = originalShape ?? new VisioShape("geometry") {
            Width = x == 0 ? 0 : shape.Width / x, Height = y == 0 ? 0 : shape.Height / y,
            LocPinX = x == 0 ? 0 : shape.LocPinX / x, LocPinY = y == 0 ? 0 : shape.LocPinY / y,
            PinX = x == 0 ? 0 : shape.PinX / x, PinY = y == 0 ? 0 : shape.PinY / y, Angle = shape.Angle
        };
        foreach (XElement section in shape.PreservedGeometrySections) {
            XNamespace ns = section.Name.Namespace;
            (double X, double Y)? previous = null;
            foreach (XElement row in section.Elements(ns + "Row")) {
                if (VisioShapeGeometry.IsDeleted(row)) continue;
                string type = (string?)row.Attribute("T") ?? "Geometry";
                if (type == "Geometry") continue;
                ValidateReadableRow(row, type, original, requireCompleteRows);
                bool hasEnd = VisioShapeGeometry.TryReadPoint(row, ns, original, type.StartsWith("Rel", StringComparison.Ordinal), out var end);
                if (!uniform && (type == "ArcTo" || type == "EllipticalArcTo" || type == "RelEllipticalArcTo")) {
                    ScaleArc(row, type, shape, original, previous, hasEnd ? end : null, x, y);
                    previous = hasEnd ? end : null;
                    continue;
                }
                previous = type == "Ellipse" || type == "InfiniteLine" || !hasEnd ? null : end;
                if (type.StartsWith("Rel", StringComparison.Ordinal)) {
                    MaterializeRelativeCells(row, original, shape);
                    continue;
                }
                if (type != "MoveTo" && type != "LineTo" && type != "Ellipse" && type != "InfiniteLine" && type != "NURBSTo" && type != "PolylineTo" &&
                    type != "SplineStart" && type != "SplineKnot" && type != "CubBezTo" && type != "QuadBezTo" &&
                    !(uniform && (type == "ArcTo" || type == "EllipticalArcTo")))
                    throw new NotSupportedException("Resizing geometry row '" + type + "' with these scale factors is not supported.");
                foreach (XElement cell in row.Elements(ns + "Cell")) {
                    string? name = (string?)cell.Attribute("N");
                    if ((type == "NURBSTo" && name == "E") || (type == "PolylineTo" && name == "A")) {
                        ScalePointFormula(cell, type == "NURBSTo", original, x, y);
                        continue;
                    }
                    double factor = name == "X" ? x : name == "Y" ? y : 1;
                    if (type == "Ellipse" || type == "CubBezTo") factor = name == "A" || name == "C" ? x : name == "B" || name == "D" ? y : factor;
                    if (type == "InfiniteLine" || type == "EllipticalArcTo" || type == "QuadBezTo") factor = name == "A" ? x : name == "B" ? y : factor;
                    if (type == "ArcTo" && name == "A") factor = x;
                    if (!VisioShapeGeometry.TryParseCellLiteral((string?)cell.Attribute("V"), original, out double value) &&
                        !VisioShapeGeometry.TryParseCellLiteral((string?)cell.Attribute("F"), original, out value)) continue;
                    MaterializeCell(cell, value * factor, shape);
                }
            }
        }
    }

    // A partial/native row can carry formulas without a numeric cache. Leaving
    // those cells live while resizing the frame would mix old and new geometry.
    // Required coordinates must be present; optional numeric profile cells must
    // be readable when present. Unknown producer metadata remains untouched.
    private static void ValidateReadableRow(XElement row, string type, VisioShape original, bool requireCompleteRows) {
        string absolute = type.StartsWith("Rel", StringComparison.Ordinal) ? type.Substring(3) : type;
        if (type.StartsWith("Rel", StringComparison.Ordinal) && absolute != "MoveTo" && absolute != "LineTo" &&
            absolute != "CubBezTo" && absolute != "QuadBezTo" && absolute != "EllipticalArcTo")
            throw new NotSupportedException("Resizing geometry row '" + type + "' is not supported.");
        string required = absolute switch {
            "MoveTo" or "LineTo" or "SplineStart" or "SplineKnot" or "PolylineTo" => "XY",
            "ArcTo" => "XYA",
            "InfiniteLine" or "QuadBezTo" or "EllipticalArcTo" => "XYAB",
            "Ellipse" or "CubBezTo" or "NURBSTo" => "XYABCD",
            _ => throw new NotSupportedException("Resizing geometry row '" + type + "' is not supported.")
        };
        foreach (char name in required) Require(name.ToString());
        if (absolute == "NURBSTo" || absolute == "PolylineTo") {
            string name = absolute == "NURBSTo" ? "E" : "A";
            if (requireCompleteRows && !row.Elements(row.Name.Namespace + "Cell").Any(cell => (string?)cell.Attribute("N") == name))
                throw new NotSupportedException("Resizing geometry row '" + type + "' requires a readable control formula.");
        }
        if (absolute == "EllipticalArcTo" || absolute == "SplineStart" || absolute == "SplineKnot") {
            foreach (XElement cell in row.Elements(row.Name.Namespace + "Cell")) {
                string? name = (string?)cell.Attribute("N");
                if (name is "A" or "B" or "C" or "D") Require(name);
            }
        }
        void Require(string name) {
            // Native connector saving retains sparse master deltas. Their absent
            // cells stay inherited; explicit resize/creation requires complete rows.
            if (!requireCompleteRows && !row.Elements(row.Name.Namespace + "Cell").Any(cell => (string?)cell.Attribute("N") == name)) return;
            if (!VisioShapeGeometry.TryReadCell(row, row.Name.Namespace, name, original, out _))
                throw new NotSupportedException("Resizing geometry row '" + type + "' requires a readable '" + name + "' cell.");
        }
    }

    // Normalized geometry values remain fractions when the shape changes size.
    // Their formulas still need validation against the final instance frame.
    private static void MaterializeRelativeCells(XElement row, VisioShape original, VisioShape shape) {
        foreach (XElement cell in row.Elements(row.Name.Namespace + "Cell")) {
            if (VisioShapeGeometry.TryParseCellLiteral((string?)cell.Attribute("V"), original, out double value) ||
                VisioShapeGeometry.TryParseCellLiteral((string?)cell.Attribute("F"), original, out value))
                MaterializeCell(cell, value, shape);
        }
    }

    private static void ScalePointFormula(XElement cell, bool nurbs, VisioShape original, double x, double y) {
        string function = nurbs ? "NURBS" : "POLYLINE";
        string? formula = (string?)cell.Attribute("F") ?? (string?)cell.Attribute("V");
        if (!VisioShapeGeometry.TryParseFunctionArguments(formula, function, out var args))
            throw new NotSupportedException("Resizing geometry requires a readable " + function + " formula.");
        int header = nurbs ? 4 : 2, stride = nurbs ? 4 : 2, typeIndex = nurbs ? 2 : 0;
        if (args.Count <= header || (args.Count - header) % stride != 0 ||
            !int.TryParse(args[typeIndex], out int xType) || !int.TryParse(args[typeIndex + 1], out int yType) ||
            (xType != 0 && xType != 1) || (yType != 0 && yType != 1))
            throw new NotSupportedException("Unsupported coordinate profile in " + function + " geometry.");
        if (nurbs) {
            ScaleCoordinate(0, 1);
            double degree = ScaleCoordinate(1, 1);
            if (degree < 1 || degree > 25 || degree != Math.Round(degree))
                throw new NotSupportedException("Resizing NURBS requires an integer degree between 1 and 25.");
        }
        for (int i = header; i < args.Count; i += stride) {
            ScaleCoordinate(i, xType == 1 ? x : 1);
            ScaleCoordinate(i + 1, yType == 1 ? y : 1);
            if (nurbs) {
                ScaleCoordinate(i + 2, 1);
                if (ScaleCoordinate(i + 3, 1) <= 0)
                    throw new NotSupportedException("Resizing NURBS requires positive control weights.");
            }
        }
        string result = function + "(" + string.Join(",", args) + ")";
        cell.SetAttributeValue("F", result);
        cell.SetAttributeValue("V", result);
        double ScaleCoordinate(int index, double factor) {
            if (!VisioShapeGeometry.TryParseCellLiteral(args[index], original, out double value))
                throw new NotSupportedException("Resizing " + function + " requires evaluable numeric arguments.");
            double transformed = value * factor;
            if (!Finite(transformed)) throw new NotSupportedException("Resizing " + function + " requires finite transformed control coordinates.");
            args[index] = transformed.ToString("R", CultureInfo.InvariantCulture);
            return transformed;
        }
    }

    /// <summary>Writes a finite resize cache and retains only a formula that reproduces it in the final frame.</summary>
    internal static void MaterializeCell(XElement cell, double value, VisioShape shape, bool replaceMeaning = false) {
        if (double.IsNaN(value) || double.IsInfinity(value))
            throw new NotSupportedException("Resizing geometry requires finite cached coordinates.");
        string cached = value.ToString("R", CultureInfo.InvariantCulture);
        bool changed = (string?)cell.Attribute("V") != cached;
        cell.SetAttributeValue("V", cached);
        // Expressions that already produce the requested coordinate remain live.
        // Repurposed cells and other expressions are materialized on the instance only.
        if (cell.Attribute("F") is XAttribute formula && (replaceMeaning ||
            !VisioShapeGeometry.TryParseCellLiteral(formula.Value, shape, out double evaluated) || Math.Abs(evaluated - value) > 1e-10)) {
            formula.Remove();
            changed = true;
        }
        if (changed || replaceMeaning) cell.Attribute("E")?.Remove();
        if (replaceMeaning) shape.NativeCellMetadata?.ForgetCellState(cell);
    }
}
