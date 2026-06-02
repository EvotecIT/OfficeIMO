using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Xml.Linq;

namespace OfficeIMO.Visio {
    internal static partial class VisioShapeGeometry {

        private static bool TryParseGeometrySection(VisioShape shape, XElement section, int fillGroup, out List<VisioShapeGeometryPath> paths) {
            paths = new List<VisioShapeGeometryPath>();
            List<(double X, double Y)> points = new();
            XNamespace ns = section.Name.Namespace;
            ReadGeometryFlags(shape, section, ns, out bool noFill, out bool noLine, out bool noShow);
            if (noShow) {
                return true;
            }

            bool handledStandaloneGeometry = false;
            bool closedPath = true;
            bool currentNoFill = noFill;
            bool currentNoLine = noLine;
            List<XElement> rows = section.Elements(ns + "Row").ToList();
            for (int rowIndex = 0; rowIndex < rows.Count; rowIndex++) {
                XElement row = rows[rowIndex];
                if (IsDeleted(row)) {
                    continue;
                }

                string? type = row.Attribute("T")?.Value;
                if (string.Equals(type, "Geometry", StringComparison.OrdinalIgnoreCase)) {
                    continue;
                }

                if (handledStandaloneGeometry) {
                    return false;
                }

                if (string.Equals(type, "Ellipse", StringComparison.OrdinalIgnoreCase)) {
                    if (points.Count != 0 ||
                        !TryReadCell(row, ns, "X", shape, out double centerX) ||
                        !TryReadCell(row, ns, "Y", shape, out double centerY) ||
                        !TryReadCell(row, ns, "A", shape, out double point1X) ||
                        !TryReadCell(row, ns, "B", shape, out double point1Y) ||
                        !TryReadCell(row, ns, "C", shape, out double point2X) ||
                        !TryReadCell(row, ns, "D", shape, out double point2Y)) {
                        return false;
                    }

                    AppendEllipse(points, (centerX, centerY), (point1X, point1Y), (point2X, point2Y));
                    handledStandaloneGeometry = true;
                    continue;
                }

                if (string.Equals(type, "InfiniteLine", StringComparison.OrdinalIgnoreCase)) {
                    if (points.Count != 0 ||
                        !TryReadCell(row, ns, "X", shape, out double point1X) ||
                        !TryReadCell(row, ns, "Y", shape, out double point1Y) ||
                        !TryReadCell(row, ns, "A", shape, out double point2X) ||
                        !TryReadCell(row, ns, "B", shape, out double point2Y) ||
                        !TryClipInfiniteLineToShapeBounds(shape, (point1X, point1Y), (point2X, point2Y), out List<(double X, double Y)> linePoints)) {
                        return false;
                    }

                    points.AddRange(linePoints);
                    closedPath = false;
                    handledStandaloneGeometry = true;
                    continue;
                }

                if (string.Equals(type, "ArcTo", StringComparison.OrdinalIgnoreCase)) {
                    if (points.Count == 0 ||
                        !TryReadPoint(row, ns, shape, relative: false, out (double X, double Y) arcEnd) ||
                        !TryReadCell(row, ns, "A", shape, out double arcHeight)) {
                        return false;
                    }

                    AppendArc(points, points[points.Count - 1], arcEnd, arcHeight);
                    continue;
                }

                if (string.Equals(type, "EllipticalArcTo", StringComparison.OrdinalIgnoreCase)) {
                    if (points.Count == 0 ||
                        !TryReadPoint(row, ns, shape, relative: false, out (double X, double Y) arcEnd) ||
                        !TryReadCell(row, ns, "A", shape, out double controlX) ||
                        !TryReadCell(row, ns, "B", shape, out double controlY)) {
                        return false;
                    }

                    double angle = TryReadCell(row, ns, "C", shape, out double rawAngle) ? rawAngle : 0D;
                    double ratio = TryReadCell(row, ns, "D", shape, out double rawRatio) ? rawRatio : 1D;
                    AppendEllipticalArc(points, points[points.Count - 1], arcEnd, (controlX, controlY), angle, ratio);
                    continue;
                }

                if (string.Equals(type, "RelEllipticalArcTo", StringComparison.OrdinalIgnoreCase)) {
                    if (points.Count == 0 ||
                        !TryReadPoint(row, ns, shape, relative: true, out (double X, double Y) arcEnd) ||
                        !TryReadRelativeCell(row, ns, "A", shape.Width, out double controlX) ||
                        !TryReadRelativeCell(row, ns, "B", shape.Height, out double controlY)) {
                        return false;
                    }

                    double angle = TryReadCell(row, ns, "C", shape, out double rawAngle) ? rawAngle : 0D;
                    double ratio = TryReadCell(row, ns, "D", shape, out double rawRatio) ? rawRatio : 1D;
                    AppendEllipticalArc(points, points[points.Count - 1], arcEnd, (controlX, controlY), angle, ratio);
                    continue;
                }

                if (string.Equals(type, "NURBSTo", StringComparison.OrdinalIgnoreCase)) {
                    if (points.Count == 0 ||
                        !TryReadPoint(row, ns, shape, relative: false, out (double X, double Y) end) ||
                        !TryReadCell(row, ns, "A", shape, out double secondLastKnot) ||
                        !TryReadCell(row, ns, "B", shape, out double lastWeight) ||
                        !TryReadCell(row, ns, "C", shape, out double firstKnot) ||
                        !TryReadCell(row, ns, "D", shape, out double firstWeight) ||
                        !TryReadFormulaCell(row, ns, "E", out string? nurbsFormula) ||
                        !TryParseNurbsFormula(nurbsFormula, shape, points[points.Count - 1], end, firstKnot, firstWeight, secondLastKnot, lastWeight, out NurbsCurve? nurbsCurve)) {
                        return false;
                    }

                    AppendNurbs(points, nurbsCurve!);
                    continue;
                }

                if (string.Equals(type, "SplineStart", StringComparison.OrdinalIgnoreCase)) {
                    if (points.Count == 0 ||
                        !TryReadPoint(row, ns, shape, relative: false, out (double X, double Y) secondControl)) {
                        return false;
                    }

                    int degree = TryReadCell(row, ns, "D", shape, out double rawDegree)
                        ? (int)Math.Round(rawDegree)
                        : 3;
                    if (degree < 1 || degree > 25) {
                        return false;
                    }

                    List<(double X, double Y)> splinePoints = new() {
                        points[points.Count - 1],
                        secondControl
                    };
                    while (rowIndex + 1 < rows.Count &&
                           string.Equals(rows[rowIndex + 1].Attribute("T")?.Value, "SplineKnot", StringComparison.OrdinalIgnoreCase)) {
                        rowIndex++;
                        if (IsDeleted(rows[rowIndex])) {
                            continue;
                        }

                        if (!TryReadPoint(rows[rowIndex], ns, shape, relative: false, out (double X, double Y) knotPoint)) {
                            return false;
                        }

                        splinePoints.Add(knotPoint);
                    }

                    AppendSpline(points, splinePoints, degree);
                    continue;
                }

                if (string.Equals(type, "CubBezTo", StringComparison.OrdinalIgnoreCase)) {
                    if (points.Count == 0 ||
                        !TryReadPoint(row, ns, shape, relative: false, out (double X, double Y) end) ||
                        !TryReadCell(row, ns, "A", shape, out double control1X) ||
                        !TryReadCell(row, ns, "B", shape, out double control1Y) ||
                        !TryReadCell(row, ns, "C", shape, out double control2X) ||
                        !TryReadCell(row, ns, "D", shape, out double control2Y)) {
                        return false;
                    }

                    AppendCubicBezier(points, points[points.Count - 1], (control1X, control1Y), (control2X, control2Y), end);
                    continue;
                }

                if (string.Equals(type, "RelCubBezTo", StringComparison.OrdinalIgnoreCase)) {
                    if (points.Count == 0 ||
                        !TryReadPoint(row, ns, shape, relative: true, out (double X, double Y) end) ||
                        !TryReadRelativeCell(row, ns, "A", shape.Width, out double control1X) ||
                        !TryReadRelativeCell(row, ns, "B", shape.Height, out double control1Y) ||
                        !TryReadRelativeCell(row, ns, "C", shape.Width, out double control2X) ||
                        !TryReadRelativeCell(row, ns, "D", shape.Height, out double control2Y)) {
                        return false;
                    }

                    AppendCubicBezier(points, points[points.Count - 1], (control1X, control1Y), (control2X, control2Y), end);
                    continue;
                }

                if (string.Equals(type, "QuadBezTo", StringComparison.OrdinalIgnoreCase)) {
                    if (points.Count == 0 ||
                        !TryReadPoint(row, ns, shape, relative: false, out (double X, double Y) end) ||
                        !TryReadCell(row, ns, "A", shape, out double controlX) ||
                        !TryReadCell(row, ns, "B", shape, out double controlY)) {
                        return false;
                    }

                    AppendQuadraticBezier(points, points[points.Count - 1], (controlX, controlY), end);
                    continue;
                }

                if (string.Equals(type, "RelQuadBezTo", StringComparison.OrdinalIgnoreCase)) {
                    if (points.Count == 0 ||
                        !TryReadPoint(row, ns, shape, relative: true, out (double X, double Y) end) ||
                        !TryReadRelativeCell(row, ns, "A", shape.Width, out double controlX) ||
                        !TryReadRelativeCell(row, ns, "B", shape.Height, out double controlY)) {
                        return false;
                    }

                    AppendQuadraticBezier(points, points[points.Count - 1], (controlX, controlY), end);
                    continue;
                }

                if (string.Equals(type, "PolylineTo", StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(type, "PolyLineTo", StringComparison.OrdinalIgnoreCase)) {
                    if (points.Count == 0 ||
                        !TryReadPoint(row, ns, shape, relative: false, out (double X, double Y) end)) {
                        return false;
                    }

                    if (TryReadFormulaCell(row, ns, "A", out string? formula) &&
                        TryParsePolylineFormula(formula, shape, out List<(double X, double Y)> polylinePoints) &&
                        polylinePoints.Count > 0) {
                        AppendPolyline(points, polylinePoints, end);
                    } else {
                        points.Add(end);
                    }

                    continue;
                }

                bool isMove;
                bool relative;
                if (string.Equals(type, "MoveTo", StringComparison.OrdinalIgnoreCase)) {
                    isMove = true;
                    relative = false;
                } else if (string.Equals(type, "LineTo", StringComparison.OrdinalIgnoreCase)) {
                    isMove = false;
                    relative = false;
                } else if (string.Equals(type, "RelMoveTo", StringComparison.OrdinalIgnoreCase)) {
                    isMove = true;
                    relative = true;
                } else if (string.Equals(type, "RelLineTo", StringComparison.OrdinalIgnoreCase)) {
                    isMove = false;
                    relative = true;
                } else {
                    return false;
                }

                if (!TryReadPoint(row, ns, shape, relative, out (double X, double Y) point)) {
                    return false;
                }

                if (isMove && points.Count > 0) {
                    if (!TryAddGeometryPath(paths, points, currentNoFill, currentNoLine, closedPath, fillGroup)) {
                        return false;
                    }

                    points = new List<(double X, double Y)>();
                    closedPath = true;
                    currentNoFill = noFill;
                    currentNoLine = noLine;
                }

                if (isMove) {
                    if (TryReadBooleanCell(row, ns, "NoFill", shape, out bool rowNoFill)) {
                        currentNoFill = rowNoFill;
                    }

                    if (TryReadBooleanCell(row, ns, "NoLine", shape, out bool rowNoLine)) {
                        currentNoLine = rowNoLine;
                    }
                }

                points.Add(point);
            }

            return TryAddGeometryPath(paths, points, currentNoFill, currentNoLine, closedPath, fillGroup);
        }

        private static bool TryAddGeometryPath(
            List<VisioShapeGeometryPath> paths,
            List<(double X, double Y)> points,
            bool noFill,
            bool noLine,
            bool closedPath,
            int fillGroup) {
            bool explicitlyClosed = points.Count > 1 && NearlyEqual(points[0], points[points.Count - 1]);
            bool renderClosedPath = closedPath && (!noFill || explicitlyClosed);
            int minimumPointCount = renderClosedPath ? 3 : 2;
            if (points.Count < minimumPointCount) {
                return false;
            }

            if (renderClosedPath && explicitlyClosed) {
                points.RemoveAt(points.Count - 1);
            }

            if (points.Count < minimumPointCount) {
                return false;
            }

            paths.Add(new VisioShapeGeometryPath(points, noFill, noLine, renderClosedPath, fillGroup));
            return true;
        }

        private static void ReadGeometryFlags(VisioShape shape, XElement section, XNamespace ns, out bool noFill, out bool noLine, out bool noShow) {
            noFill = false;
            noLine = false;
            noShow = false;
            foreach (XElement row in section.Elements(ns + "Row")) {
                if (IsDeleted(row)) {
                    continue;
                }

                if (!string.Equals(row.Attribute("T")?.Value, "Geometry", StringComparison.OrdinalIgnoreCase)) {
                    continue;
                }

                if (TryReadBooleanCell(row, ns, "NoFill", shape, out bool rowNoFill)) {
                    noFill = rowNoFill;
                }

                if (TryReadBooleanCell(row, ns, "NoLine", shape, out bool rowNoLine)) {
                    noLine = rowNoLine;
                }

                if (TryReadBooleanCell(row, ns, "NoShow", shape, out bool rowNoShow)) {
                    noShow = rowNoShow;
                }
            }
        }

        private static bool IsDeleted(XElement row) {
            string? raw = row.Attribute("Del")?.Value;
            return string.Equals(raw, "1", StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(raw, "true", StringComparison.OrdinalIgnoreCase);
        }
    }
}
