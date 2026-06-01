using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Xml.Linq;

namespace OfficeIMO.Visio {
    internal sealed class VisioShapeGeometryPath {
        internal VisioShapeGeometryPath(List<(double X, double Y)> points, bool noFill, bool noLine, bool isClosed, int fillGroup) {
            Points = points;
            NoFill = noFill;
            NoLine = noLine;
            IsClosed = isClosed;
            FillGroup = fillGroup;
        }

        internal List<(double X, double Y)> Points { get; }

        internal bool NoFill { get; }

        internal bool NoLine { get; }

        internal bool IsClosed { get; }

        internal int FillGroup { get; }
    }

    internal static class VisioShapeGeometry {
        private const int ArcSegmentCount = 16;

        internal static string ResolveRenderKind(VisioShape shape) {
            string kind = NormalizeKind(shape.MasterNameU ?? shape.NameU ?? shape.Name ?? string.Empty);
            if (IsSemanticTerminatorShape(shape, kind)) {
                return "terminator";
            }

            if (IsSemanticDocumentShape(shape, kind)) {
                return "document";
            }

            if (IsSemanticDatabaseShape(shape, kind)) {
                return "database";
            }

            return kind;
        }

        internal static bool TryGetRenderClosedPaths(VisioShape shape, out List<VisioShapeGeometryPath> paths) {
            if (TryGetPreservedClosedPaths(shape, out paths)) {
                return true;
            }

            VisioShape? masterShape = shape.MasterShape ?? shape.Master?.Shape;
            if (masterShape == null ||
                !TryGetPreservedClosedPaths(masterShape, out List<VisioShapeGeometryPath> masterPaths)) {
                paths = new List<VisioShapeGeometryPath>();
                return false;
            }

            paths = ScalePaths(masterShape, shape, masterPaths);
            return true;
        }

        internal static bool TryGetPreservedClosedPaths(VisioShape shape, out List<VisioShapeGeometryPath> paths) {
            paths = new List<VisioShapeGeometryPath>();
            bool handledGeometry = false;
            int fillGroup = 0;
            foreach (XElement section in shape.PreservedGeometrySections) {
                if (!TryParseGeometrySection(shape, section, fillGroup, out List<VisioShapeGeometryPath> sectionPaths)) {
                    continue;
                }

                handledGeometry = true;
                paths.AddRange(sectionPaths);
                fillGroup++;
            }

            return handledGeometry;
        }

        private static List<VisioShapeGeometryPath> ScalePaths(VisioShape source, VisioShape target, List<VisioShapeGeometryPath> sourcePaths) {
            List<VisioShapeGeometryPath> scaledPaths = new();
            if (source.Width <= 0D || source.Height <= 0D || target.Width <= 0D || target.Height <= 0D) {
                return scaledPaths;
            }

            double scaleX = target.Width / source.Width;
            double scaleY = target.Height / source.Height;
            foreach (VisioShapeGeometryPath sourcePath in sourcePaths) {
                List<(double X, double Y)> scaledPath = new();
                for (int i = 0; i < sourcePath.Points.Count; i++) {
                    scaledPath.Add((sourcePath.Points[i].X * scaleX, sourcePath.Points[i].Y * scaleY));
                }

                int minimumPoints = sourcePath.IsClosed ? 3 : 2;
                if (scaledPath.Count >= minimumPoints) {
                    scaledPaths.Add(new VisioShapeGeometryPath(scaledPath, sourcePath.NoFill, sourcePath.NoLine, sourcePath.IsClosed, sourcePath.FillGroup));
                }
            }

            return scaledPaths;
        }

        internal static List<(double X, double Y)> GetBuiltinClosedPath(VisioShape shape, string kind) {
            double width = shape.Width;
            double height = shape.Height;
            double midX = width / 2D;
            double midY = height / 2D;
            switch (kind) {
                case "diamond":
                case "decision":
                    return new List<(double X, double Y)> { (midX, 0), (width, midY), (midX, height), (0, midY) };
                case "terminator":
                case "startend":
                    return CreateTerminatorPath(width, height);
                case "document":
                    return CreateDocumentPath(width, height);
                case "delay":
                    return CreateDelayPath(width, height);
                case "triangle":
                    return new List<(double X, double Y)> { (0, 0), (midX, height), (width, 0) };
                case "pentagon":
                case "offpagereference":
                    return new List<(double X, double Y)> {
                        (midX, height),
                        (width, height * 0.62D),
                        (width * 0.8D, 0),
                        (width * 0.2D, 0),
                        (0, height * 0.62D)
                    };
                case "parallelogram":
                case "data":
                    double offset = Math.Min(width / 4D, Math.Max(width / 10D, height / 3D));
                    return new List<(double X, double Y)> { (offset, 0), (width, 0), (width - offset, height), (0, height) };
                case "hexagon":
                case "preparation":
                    double inset = Math.Min(width / 4D, Math.Max(width / 8D, height / 4D));
                    return new List<(double X, double Y)> { (inset, 0), (width - inset, 0), (width, midY), (width - inset, height), (inset, height), (0, midY) };
                case "chevron":
                    double chevronInset = width * 0.28D;
                    double chevronPointBase = width - chevronInset;
                    return new List<(double X, double Y)> {
                        (0, 0),
                        (chevronPointBase, 0),
                        (width, midY),
                        (chevronPointBase, height),
                        (0, height),
                        (chevronInset, midY)
                    };
                case "trapezoid":
                case "manualoperation":
                    double trapInset = Math.Min(width / 5D, Math.Max(width / 10D, height / 4D));
                    return new List<(double X, double Y)> { (trapInset, height), (width - trapInset, height), (width, 0), (0, 0) };
                case "manualinput":
                    double manualInputTopRightY = height * 0.75D;
                    return new List<(double X, double Y)> { (0, 0), (width, 0), (width, manualInputTopRightY), (0, height) };
                default:
                    return new List<(double X, double Y)> { (0, 0), (width, 0), (width, height), (0, height) };
            }
        }

        private static List<(double X, double Y)> CreateTerminatorPath(double width, double height) {
            if (width <= 0D || height <= 0D) {
                return new List<(double X, double Y)> { (0, 0), (width, 0), (width, height), (0, height) };
            }

            double radius = Math.Min(width, height) / 2D;
            double leftCenterX = radius;
            double rightCenterX = width - radius;
            double centerY = height / 2D;
            List<(double X, double Y)> points = new() {
                (leftCenterX, 0D),
                (rightCenterX, 0D)
            };

            AppendArcPoints(points, rightCenterX, centerY, radius, -Math.PI / 2D, Math.PI / 2D);
            points.Add((leftCenterX, height));
            AppendArcPoints(points, leftCenterX, centerY, radius, Math.PI / 2D, Math.PI * 1.5D);
            return points;
        }

        private static List<(double X, double Y)> CreateDocumentPath(double width, double height) {
            if (width <= 0D || height <= 0D) {
                return new List<(double X, double Y)> { (0, 0), (width, 0), (width, height), (0, height) };
            }

            double waveDepth = Math.Min(height * 0.16D, width * 0.1D);
            double baseY = waveDepth;
            double crestY = Math.Min(height, baseY + (waveDepth * 0.55D));
            double troughY = Math.Max(0D, baseY - (waveDepth * 0.55D));
            return new List<(double X, double Y)> {
                (0D, height),
                (width, height),
                (width, baseY),
                (width * 0.86D, troughY),
                (width * 0.72D, baseY),
                (width * 0.58D, crestY),
                (width * 0.44D, baseY),
                (width * 0.30D, troughY),
                (width * 0.15D, baseY),
                (0D, baseY)
            };
        }

        private static List<(double X, double Y)> CreateDelayPath(double width, double height) {
            if (width <= 0D || height <= 0D) {
                return new List<(double X, double Y)> { (0, 0), (width, 0), (width, height), (0, height) };
            }

            double radius = Math.Min(width, height) / 2D;
            double rightCenterX = width - radius;
            double centerY = height / 2D;
            List<(double X, double Y)> points = new() {
                (0D, 0D),
                (rightCenterX, 0D)
            };

            AppendArcPoints(points, rightCenterX, centerY, radius, -Math.PI / 2D, Math.PI / 2D);
            points.Add((0D, height));
            return points;
        }

        private static void AppendArcPoints(List<(double X, double Y)> points, double centerX, double centerY, double radius, double startAngle, double endAngle) {
            for (int i = 1; i <= ArcSegmentCount; i++) {
                double t = i / (double)ArcSegmentCount;
                double angle = startAngle + ((endAngle - startAngle) * t);
                points.Add((centerX + Math.Cos(angle) * radius, centerY + Math.Sin(angle) * radius));
            }
        }

        private static bool IsSemanticTerminatorShape(VisioShape shape, string kind) {
            if (kind == "terminator" || kind == "startend") {
                return true;
            }

            if (kind != "ellipse" && kind != "circle" && kind != "process" && kind != "rectangle") {
                return false;
            }

            string metadata = GetStencilMetadata(shape);
            return ContainsSemanticTerm(metadata, "terminator") ||
                   ContainsSemanticTerm(metadata, "startend") ||
                   ContainsSemanticTerm(metadata, "flowstartend") ||
                   ContainsSemanticTerm(metadata, "swimstartend");
        }

        private static bool IsSemanticDocumentShape(VisioShape shape, string kind) {
            if (kind == "document") {
                return true;
            }

            if (kind != "data" && kind != "process" && kind != "rectangle") {
                return false;
            }

            string identity = GetStencilIdentityMetadata(shape);
            return ContainsSemanticTerm(identity, "document") ||
                   ContainsSemanticTerm(identity, "businessdocument") ||
                   ContainsSemanticTerm(identity, "collabdocument");
        }

        private static bool IsSemanticDatabaseShape(VisioShape shape, string kind) {
            if (kind == "database" || kind == "datastore" || kind == "storage") {
                return true;
            }

            if (kind != "data" && kind != "process" && kind != "rectangle") {
                return false;
            }

            string metadata = GetStencilMetadata(shape);
            return ContainsSemanticTerm(metadata, "database") ||
                   ContainsSemanticTerm(metadata, "datastore") ||
                   ContainsSemanticTerm(metadata, "storage") ||
                   ContainsSemanticTerm(metadata, "warehouse") ||
                   ContainsSemanticTerm(metadata, "datalake") ||
                   ContainsSemanticTerm(metadata, "sql");
        }

        private static string GetStencilIdentityMetadata(VisioShape shape) {
            VisioMaster? master = shape.Master;
            return NormalizeKind(
                (shape.GetUserCellValue(VisioSemanticUserCells.StencilId) ?? string.Empty) + " " +
                (shape.GetUserCellValue(VisioSemanticUserCells.StencilName) ?? string.Empty) + " " +
                (master?.StencilId ?? string.Empty) + " " +
                (master?.StencilName ?? string.Empty));
        }

        private static string GetStencilMetadata(VisioShape shape) {
            VisioMaster? master = shape.Master;
            return NormalizeKind(
                (shape.GetUserCellValue(VisioSemanticUserCells.StencilId) ?? string.Empty) + " " +
                (shape.GetUserCellValue(VisioSemanticUserCells.StencilName) ?? string.Empty) + " " +
                (shape.GetUserCellValue(VisioSemanticUserCells.StencilCategory) ?? string.Empty) + " " +
                (shape.GetUserCellValue(VisioSemanticUserCells.StencilAliases) ?? string.Empty) + " " +
                (shape.GetUserCellValue(VisioSemanticUserCells.StencilTags) ?? string.Empty) + " " +
                (master?.StencilId ?? string.Empty) + " " +
                (master?.StencilName ?? string.Empty) + " " +
                (master?.StencilCategory ?? string.Empty) + " " +
                string.Join(" ", master?.StencilAliases ?? Array.Empty<string>()) + " " +
                string.Join(" ", master?.StencilTags ?? Array.Empty<string>()));
        }

        private static bool ContainsSemanticTerm(string metadata, string term) =>
            metadata.IndexOf(term, StringComparison.OrdinalIgnoreCase) >= 0;

        private static string NormalizeKind(string value) {
            if (string.IsNullOrWhiteSpace(value)) {
                return string.Empty;
            }

            char[] buffer = new char[value.Length];
            int index = 0;
            foreach (char c in value) {
                if (char.IsLetterOrDigit(c)) {
                    buffer[index++] = char.ToLowerInvariant(c);
                }
            }

            return new string(buffer, 0, index);
        }

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
                    if (!TryAddGeometryPath(paths, points, noFill, noLine, closedPath, fillGroup)) {
                        return false;
                    }

                    points = new List<(double X, double Y)>();
                    closedPath = true;
                }

                points.Add(point);
            }

            return TryAddGeometryPath(paths, points, noFill, noLine, closedPath, fillGroup);
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

        private static bool TryReadPoint(XElement row, XNamespace ns, VisioShape shape, bool relative, out (double X, double Y) point) {
            if (!TryReadCell(row, ns, "X", shape, out double x) ||
                !TryReadCell(row, ns, "Y", shape, out double y)) {
                point = default;
                return false;
            }

            point = relative ? (x * shape.Width, y * shape.Height) : (x, y);
            return true;
        }

        private static bool TryReadRelativeCell(XElement row, XNamespace ns, string name, double scale, out double value) {
            if (!TryReadRawCell(row, ns, name, out double raw)) {
                value = 0D;
                return false;
            }

            value = raw * scale;
            return true;
        }

        private static void AppendArc(List<(double X, double Y)> path, (double X, double Y) start, (double X, double Y) end, double sagitta) {
            double dx = end.X - start.X;
            double dy = end.Y - start.Y;
            double chord = Math.Sqrt((dx * dx) + (dy * dy));
            if (chord <= 1e-9 || Math.Abs(sagitta) <= 1e-9) {
                path.Add(end);
                return;
            }

            double midX = (start.X + end.X) / 2D;
            double midY = (start.Y + end.Y) / 2D;
            double normalX = -dy / chord;
            double normalY = dx / chord;
            double radius = ((chord * chord) / (8D * Math.Abs(sagitta))) + (Math.Abs(sagitta) / 2D);
            double signedCenterOffset = Math.Sign(sagitta) * (radius - Math.Abs(sagitta));
            double centerX = midX - (normalX * signedCenterOffset);
            double centerY = midY - (normalY * signedCenterOffset);
            double arcPointX = midX + (normalX * sagitta);
            double arcPointY = midY + (normalY * sagitta);
            double startAngle = Math.Atan2(start.Y - centerY, start.X - centerX);
            double endAngle = Math.Atan2(end.Y - centerY, end.X - centerX);
            double midAngle = Math.Atan2(arcPointY - centerY, arcPointX - centerX);
            double sweep = NormalizeSweep(endAngle - startAngle);
            if (!AngleLiesOnSweep(startAngle, sweep, midAngle)) {
                sweep = sweep > 0D ? sweep - (Math.PI * 2D) : sweep + (Math.PI * 2D);
            }

            for (int i = 1; i <= ArcSegmentCount; i++) {
                double angle = startAngle + (sweep * i / ArcSegmentCount);
                path.Add((centerX + (Math.Cos(angle) * radius), centerY + (Math.Sin(angle) * radius)));
            }
        }

        private static void AppendEllipticalArc(
            List<(double X, double Y)> path,
            (double X, double Y) start,
            (double X, double Y) end,
            (double X, double Y) control,
            double angle,
            double ratio) {
            if (!IsFinite(ratio) || ratio <= 0D || ratio > 1000D) {
                path.Add(end);
                return;
            }

            double cos = Math.Cos(-angle);
            double sin = Math.Sin(-angle);
            (double X, double Y) transformedStart = TransformEllipsePoint(start, cos, sin, ratio);
            (double X, double Y) transformedEnd = TransformEllipsePoint(end, cos, sin, ratio);
            (double X, double Y) transformedControl = TransformEllipsePoint(control, cos, sin, ratio);
            if (!TryGetCircumcircle(transformedStart, transformedControl, transformedEnd, out (double X, double Y) center, out double radius)) {
                path.Add(end);
                return;
            }

            double startAngle = Math.Atan2(transformedStart.Y - center.Y, transformedStart.X - center.X);
            double endAngle = Math.Atan2(transformedEnd.Y - center.Y, transformedEnd.X - center.X);
            double controlAngle = Math.Atan2(transformedControl.Y - center.Y, transformedControl.X - center.X);
            double sweep = NormalizeSweep(endAngle - startAngle);
            if (!AngleLiesOnSweep(startAngle, sweep, controlAngle)) {
                sweep = sweep > 0D ? sweep - (Math.PI * 2D) : sweep + (Math.PI * 2D);
            }

            double inverseCos = Math.Cos(angle);
            double inverseSin = Math.Sin(angle);
            for (int i = 1; i <= ArcSegmentCount; i++) {
                double currentAngle = startAngle + (sweep * i / ArcSegmentCount);
                double x = center.X + (Math.Cos(currentAngle) * radius);
                double y = center.Y + (Math.Sin(currentAngle) * radius);
                path.Add(InverseTransformEllipsePoint((x, y), inverseCos, inverseSin, ratio));
            }
        }

        private static void AppendCubicBezier(
            List<(double X, double Y)> path,
            (double X, double Y) start,
            (double X, double Y) control1,
            (double X, double Y) control2,
            (double X, double Y) end) {
            for (int i = 1; i <= ArcSegmentCount; i++) {
                double t = i / (double)ArcSegmentCount;
                double inverse = 1D - t;
                double inverseSquared = inverse * inverse;
                double tSquared = t * t;
                double x = (inverseSquared * inverse * start.X) +
                           (3D * inverseSquared * t * control1.X) +
                           (3D * inverse * tSquared * control2.X) +
                           (tSquared * t * end.X);
                double y = (inverseSquared * inverse * start.Y) +
                           (3D * inverseSquared * t * control1.Y) +
                           (3D * inverse * tSquared * control2.Y) +
                           (tSquared * t * end.Y);
                path.Add((x, y));
            }
        }

        private static void AppendQuadraticBezier(
            List<(double X, double Y)> path,
            (double X, double Y) start,
            (double X, double Y) control,
            (double X, double Y) end) {
            for (int i = 1; i <= ArcSegmentCount; i++) {
                double t = i / (double)ArcSegmentCount;
                double inverse = 1D - t;
                double x = (inverse * inverse * start.X) +
                           (2D * inverse * t * control.X) +
                           (t * t * end.X);
                double y = (inverse * inverse * start.Y) +
                           (2D * inverse * t * control.Y) +
                           (t * t * end.Y);
                path.Add((x, y));
            }
        }

        private static void AppendSpline(
            List<(double X, double Y)> path,
            List<(double X, double Y)> splinePoints,
            int degree) {
            if (splinePoints.Count <= 1) {
                return;
            }

            if (degree <= 1 || splinePoints.Count == 2) {
                for (int i = 1; i < splinePoints.Count; i++) {
                    if (!NearlyEqual(path[path.Count - 1], splinePoints[i])) {
                        path.Add(splinePoints[i]);
                    }
                }

                return;
            }

            for (int i = 0; i < splinePoints.Count - 1; i++) {
                (double X, double Y) p0 = splinePoints[Math.Max(0, i - 1)];
                (double X, double Y) p1 = splinePoints[i];
                (double X, double Y) p2 = splinePoints[i + 1];
                (double X, double Y) p3 = splinePoints[Math.Min(splinePoints.Count - 1, i + 2)];
                for (int segment = 1; segment <= ArcSegmentCount; segment++) {
                    double t = segment / (double)ArcSegmentCount;
                    double tSquared = t * t;
                    double tCubed = tSquared * t;
                    double x = 0.5D * ((2D * p1.X) +
                                       ((-p0.X + p2.X) * t) +
                                       (((2D * p0.X) - (5D * p1.X) + (4D * p2.X) - p3.X) * tSquared) +
                                       ((-p0.X + (3D * p1.X) - (3D * p2.X) + p3.X) * tCubed));
                    double y = 0.5D * ((2D * p1.Y) +
                                       ((-p0.Y + p2.Y) * t) +
                                       (((2D * p0.Y) - (5D * p1.Y) + (4D * p2.Y) - p3.Y) * tSquared) +
                                       ((-p0.Y + (3D * p1.Y) - (3D * p2.Y) + p3.Y) * tCubed));
                    (double X, double Y) point = (x, y);
                    if (!NearlyEqual(path[path.Count - 1], point)) {
                        path.Add(point);
                    }
                }
            }
        }

        private static void AppendNurbs(List<(double X, double Y)> path, NurbsCurve curve) {
            if (curve.ControlPoints.Count <= 1) {
                return;
            }

            double start = curve.Knots[curve.Degree];
            double end = curve.Knots[curve.ControlPoints.Count];
            if (!IsFinite(start) || !IsFinite(end) || end <= start) {
                path.Add(curve.ControlPoints[curve.ControlPoints.Count - 1]);
                return;
            }

            for (int spanIndex = curve.Degree; spanIndex < curve.ControlPoints.Count; spanIndex++) {
                double spanStart = curve.Knots[spanIndex];
                double spanEnd = curve.Knots[spanIndex + 1];
                if (!IsFinite(spanStart) || !IsFinite(spanEnd) || spanEnd <= spanStart) {
                    continue;
                }

                for (int segment = 1; segment <= ArcSegmentCount; segment++) {
                    double t = spanIndex == curve.ControlPoints.Count - 1 && segment == ArcSegmentCount
                        ? end
                        : spanStart + ((spanEnd - spanStart) * segment / ArcSegmentCount);
                    (double X, double Y) point = EvaluateNurbs(curve, t);
                    if (!NearlyEqual(path[path.Count - 1], point)) {
                        path.Add(point);
                    }
                }
            }
        }

        private static (double X, double Y) EvaluateNurbs(NurbsCurve curve, double t) {
            double weightedX = 0D;
            double weightedY = 0D;
            double denominator = 0D;
            for (int i = 0; i < curve.ControlPoints.Count; i++) {
                double basis = EvaluateBasis(i, curve.Degree, t, curve.Knots, curve.ControlPoints.Count);
                double weightedBasis = basis * curve.Weights[i];
                weightedX += curve.ControlPoints[i].X * weightedBasis;
                weightedY += curve.ControlPoints[i].Y * weightedBasis;
                denominator += weightedBasis;
            }

            if (Math.Abs(denominator) <= 1e-12 || !IsFinite(denominator)) {
                return curve.ControlPoints[curve.ControlPoints.Count - 1];
            }

            return (weightedX / denominator, weightedY / denominator);
        }

        private static double EvaluateBasis(int index, int degree, double t, List<double> knots, int controlPointCount) {
            if (degree == 0) {
                bool insideSpan = knots[index] <= t && t < knots[index + 1];
                bool lastKnot = Math.Abs(t - knots[controlPointCount]) <= 1e-9 && index == controlPointCount - 1;
                return insideSpan || lastKnot ? 1D : 0D;
            }

            double value = 0D;
            double leftDenominator = knots[index + degree] - knots[index];
            if (Math.Abs(leftDenominator) > 1e-12) {
                value += ((t - knots[index]) / leftDenominator) * EvaluateBasis(index, degree - 1, t, knots, controlPointCount);
            }

            double rightDenominator = knots[index + degree + 1] - knots[index + 1];
            if (Math.Abs(rightDenominator) > 1e-12) {
                value += ((knots[index + degree + 1] - t) / rightDenominator) * EvaluateBasis(index + 1, degree - 1, t, knots, controlPointCount);
            }

            return value;
        }

        private static void AppendEllipse(
            List<(double X, double Y)> path,
            (double X, double Y) center,
            (double X, double Y) point1,
            (double X, double Y) point2) {
            double axis1X = point1.X - center.X;
            double axis1Y = point1.Y - center.Y;
            double axis2X = point2.X - center.X;
            double axis2Y = point2.Y - center.Y;
            double axis1Length = Math.Sqrt((axis1X * axis1X) + (axis1Y * axis1Y));
            double axis2Length = Math.Sqrt((axis2X * axis2X) + (axis2Y * axis2Y));
            if (axis1Length <= 1e-9 || axis2Length <= 1e-9) {
                return;
            }

            int segments = ArcSegmentCount * 2;
            for (int i = 0; i < segments; i++) {
                double angle = Math.PI * 2D * i / segments;
                path.Add((
                    center.X + (Math.Cos(angle) * axis1X) + (Math.Sin(angle) * axis2X),
                    center.Y + (Math.Cos(angle) * axis1Y) + (Math.Sin(angle) * axis2Y)));
            }
        }

        private static void AppendPolyline(
            List<(double X, double Y)> path,
            List<(double X, double Y)> polylinePoints,
            (double X, double Y) end) {
            foreach ((double X, double Y) point in polylinePoints) {
                if (path.Count == 0 || !NearlyEqual(path[path.Count - 1], point)) {
                    path.Add(point);
                }
            }

            if (path.Count == 0 || !NearlyEqual(path[path.Count - 1], end)) {
                path.Add(end);
            }
        }

        private static bool TryClipInfiniteLineToShapeBounds(
            VisioShape shape,
            (double X, double Y) first,
            (double X, double Y) second,
            out List<(double X, double Y)> points) {
            points = new List<(double X, double Y)>();
            double dx = second.X - first.X;
            double dy = second.Y - first.Y;
            if (Math.Abs(dx) <= 1e-9 && Math.Abs(dy) <= 1e-9) {
                return false;
            }

            AddInfiniteLineIntersection(points, first, dx, dy, x: 0D, y: null, width: shape.Width, height: shape.Height);
            AddInfiniteLineIntersection(points, first, dx, dy, x: shape.Width, y: null, width: shape.Width, height: shape.Height);
            AddInfiniteLineIntersection(points, first, dx, dy, x: null, y: 0D, width: shape.Width, height: shape.Height);
            AddInfiniteLineIntersection(points, first, dx, dy, x: null, y: shape.Height, width: shape.Width, height: shape.Height);

            if (points.Count < 2) {
                return false;
            }

            if (points.Count > 2) {
                double bestDistance = -1D;
                (double X, double Y) bestFirst = points[0];
                (double X, double Y) bestSecond = points[1];
                for (int i = 0; i < points.Count - 1; i++) {
                    for (int j = i + 1; j < points.Count; j++) {
                        double distance = DistanceSquared(points[i], points[j]);
                        if (distance > bestDistance) {
                            bestDistance = distance;
                            bestFirst = points[i];
                            bestSecond = points[j];
                        }
                    }
                }

                points.Clear();
                points.Add(bestFirst);
                points.Add(bestSecond);
            }

            return true;
        }

        private static void AddInfiniteLineIntersection(
            List<(double X, double Y)> points,
            (double X, double Y) origin,
            double dx,
            double dy,
            double? x,
            double? y,
            double width,
            double height) {
            double t;
            double candidateX;
            double candidateY;
            if (x.HasValue) {
                if (Math.Abs(dx) <= 1e-9) {
                    return;
                }

                t = (x.Value - origin.X) / dx;
                candidateX = x.Value;
                candidateY = origin.Y + (dy * t);
            } else if (y.HasValue) {
                if (Math.Abs(dy) <= 1e-9) {
                    return;
                }

                t = (y.Value - origin.Y) / dy;
                candidateX = origin.X + (dx * t);
                candidateY = y.Value;
            } else {
                return;
            }

            const double tolerance = 1e-8;
            if (candidateX < -tolerance || candidateX > width + tolerance ||
                candidateY < -tolerance || candidateY > height + tolerance) {
                return;
            }

            (double X, double Y) point = (Clamp(candidateX, 0D, width), Clamp(candidateY, 0D, height));
            if (!points.Any(existing => NearlyEqual(existing, point))) {
                points.Add(point);
            }
        }

        private static double DistanceSquared((double X, double Y) first, (double X, double Y) second) {
            double dx = second.X - first.X;
            double dy = second.Y - first.Y;
            return (dx * dx) + (dy * dy);
        }

        private static double Clamp(double value, double minimum, double maximum) {
            if (value < minimum) {
                return minimum;
            }

            return value > maximum ? maximum : value;
        }

        private static (double X, double Y) TransformEllipsePoint((double X, double Y) point, double cos, double sin, double ratio) {
            double x = (point.X * cos) - (point.Y * sin);
            double y = (point.X * sin) + (point.Y * cos);
            return (x / ratio, y);
        }

        private static (double X, double Y) InverseTransformEllipsePoint((double X, double Y) point, double cos, double sin, double ratio) {
            double scaledX = point.X * ratio;
            return ((scaledX * cos) - (point.Y * sin), (scaledX * sin) + (point.Y * cos));
        }

        private static bool TryGetCircumcircle(
            (double X, double Y) a,
            (double X, double Y) b,
            (double X, double Y) c,
            out (double X, double Y) center,
            out double radius) {
            double d = 2D * (a.X * (b.Y - c.Y) + b.X * (c.Y - a.Y) + c.X * (a.Y - b.Y));
            if (Math.Abs(d) <= 1e-9) {
                center = default;
                radius = 0D;
                return false;
            }

            double aSquared = (a.X * a.X) + (a.Y * a.Y);
            double bSquared = (b.X * b.X) + (b.Y * b.Y);
            double cSquared = (c.X * c.X) + (c.Y * c.Y);
            center = (
                (aSquared * (b.Y - c.Y) + bSquared * (c.Y - a.Y) + cSquared * (a.Y - b.Y)) / d,
                (aSquared * (c.X - b.X) + bSquared * (a.X - c.X) + cSquared * (b.X - a.X)) / d);
            radius = Math.Sqrt(((a.X - center.X) * (a.X - center.X)) + ((a.Y - center.Y) * (a.Y - center.Y)));
            return IsFinite(radius) && radius > 1e-9;
        }

        private static double NormalizeSweep(double sweep) {
            while (sweep <= -Math.PI) {
                sweep += Math.PI * 2D;
            }

            while (sweep > Math.PI) {
                sweep -= Math.PI * 2D;
            }

            return sweep;
        }

        private static bool AngleLiesOnSweep(double startAngle, double sweep, double testAngle) {
            double relative = NormalizePositive(testAngle - startAngle);
            if (sweep >= 0D) {
                return relative <= sweep + 1e-9;
            }

            return relative >= (Math.PI * 2D) + sweep - 1e-9;
        }

        private static double NormalizePositive(double angle) {
            while (angle < 0D) {
                angle += Math.PI * 2D;
            }

            while (angle >= Math.PI * 2D) {
                angle -= Math.PI * 2D;
            }

            return angle;
        }

        private static bool TryReadCell(XElement row, XNamespace ns, string name, VisioShape shape, out double value) {
            XElement? cell = row.Elements(ns + "Cell")
                .FirstOrDefault(item => string.Equals(item.Attribute("N")?.Value, name, StringComparison.OrdinalIgnoreCase));
            if (cell == null) {
                value = 0D;
                return false;
            }

            if (TryParseCellLiteral(cell.Attribute("V")?.Value, shape, out value)) {
                return true;
            }

            if (TryParseCellLiteral(cell.Attribute("F")?.Value, shape, out value)) {
                return true;
            }

            value = 0D;
            return false;
        }

        private static bool TryReadRawCell(XElement row, XNamespace ns, string name, out double value) {
            XElement? cell = row.Elements(ns + "Cell")
                .FirstOrDefault(item => string.Equals(item.Attribute("N")?.Value, name, StringComparison.OrdinalIgnoreCase));
            if (cell == null) {
                value = 0D;
                return false;
            }

            if (TryParseLiteralWithoutShape(cell.Attribute("V")?.Value, out value)) {
                return true;
            }

            return TryParseLiteralWithoutShape(cell.Attribute("F")?.Value, out value);
        }

        private static bool TryReadFormulaCell(XElement row, XNamespace ns, string name, out string? value) {
            XElement? cell = row.Elements(ns + "Cell")
                .FirstOrDefault(item => string.Equals(item.Attribute("N")?.Value, name, StringComparison.OrdinalIgnoreCase));
            value = cell?.Attribute("F")?.Value ?? cell?.Attribute("V")?.Value;
            return !string.IsNullOrWhiteSpace(value);
        }

        private static bool TryReadBooleanCell(XElement row, XNamespace ns, string name, VisioShape shape, out bool value) {
            if (TryReadCell(row, ns, name, shape, out double numeric)) {
                value = Math.Abs(numeric) > 1e-9;
                return true;
            }

            value = false;
            return false;
        }

        private static bool TryParseCellLiteral(string? raw, VisioShape shape, out double value) {
            raw = NormalizeCellLiteral(raw);
            if (double.TryParse(raw, NumberStyles.Float, CultureInfo.InvariantCulture, out value) &&
                IsFinite(value)) {
                return true;
            }

            if (!string.IsNullOrWhiteSpace(raw) &&
                CellExpressionParser.TryEvaluate(raw!, shape.Width, shape.Height, shape.LocPinX, shape.LocPinY, shape.PinX, shape.PinY, shape.Angle, out value) &&
                IsFinite(value)) {
                return true;
            }

            value = 0D;
            return false;
        }

        private static bool TryParseLiteralWithoutShape(string? raw, out double value) {
            raw = NormalizeCellLiteral(raw);
            if (double.TryParse(raw, NumberStyles.Float, CultureInfo.InvariantCulture, out value) &&
                IsFinite(value)) {
                return true;
            }

            value = 0D;
            return false;
        }

        private static bool TryParsePolylineFormula(string? raw, VisioShape shape, out List<(double X, double Y)> points) {
            points = new List<(double X, double Y)>();
            if (!TryParseFunctionArguments(raw, "POLYLINE", out List<string> parts)) {
                return false;
            }

            if (parts.Count < 4 || parts.Count % 2 != 0) {
                return false;
            }

            if (!TryParsePolylineArgument(parts[0], shape, out double rawXType) ||
                !TryParsePolylineArgument(parts[1], shape, out double rawYType)) {
                return false;
            }

            bool xIsLocal = Math.Abs(rawXType) > 1e-9;
            bool yIsLocal = Math.Abs(rawYType) > 1e-9;
            for (int i = 2; i < parts.Count; i += 2) {
                if (!TryParsePolylineArgument(parts[i], shape, out double rawX) ||
                    !TryParsePolylineArgument(parts[i + 1], shape, out double rawY)) {
                    return false;
                }

                points.Add((xIsLocal ? rawX : rawX * shape.Width, yIsLocal ? rawY : rawY * shape.Height));
            }

            return points.Count > 0;
        }

        private static bool TryParsePolylineArgument(string raw, VisioShape shape, out double value) =>
            TryParseCellLiteral(raw, shape, out value);

        private static bool TryParseNurbsFormula(
            string? raw,
            VisioShape shape,
            (double X, double Y) start,
            (double X, double Y) end,
            double firstKnot,
            double firstWeight,
            double secondLastKnot,
            double lastWeight,
            out NurbsCurve? curve) {
            curve = null;
            if (!TryParseFunctionArguments(raw, "NURBS", out List<string> arguments) ||
                arguments.Count < 8 ||
                (arguments.Count - 4) % 4 != 0 ||
                !TryParsePolylineArgument(arguments[0], shape, out double lastKnot) ||
                !TryParsePolylineArgument(arguments[1], shape, out double rawDegree) ||
                !TryParsePolylineArgument(arguments[2], shape, out double rawXType) ||
                !TryParsePolylineArgument(arguments[3], shape, out double rawYType)) {
                return false;
            }

            int degree = (int)Math.Round(rawDegree);
            if (degree < 1 || degree > 25) {
                return false;
            }

            bool xIsLocal = Math.Abs(rawXType) > 1e-9;
            bool yIsLocal = Math.Abs(rawYType) > 1e-9;
            List<(double X, double Y)> controlPoints = new() { start };
            List<double> weights = new() { firstWeight };
            List<double> suppliedKnots = new() { firstKnot };
            for (int i = 4; i < arguments.Count; i += 4) {
                if (!TryParsePolylineArgument(arguments[i], shape, out double rawX) ||
                    !TryParsePolylineArgument(arguments[i + 1], shape, out double rawY) ||
                    !TryParsePolylineArgument(arguments[i + 2], shape, out double knot) ||
                    !TryParsePolylineArgument(arguments[i + 3], shape, out double weight)) {
                    return false;
                }

                controlPoints.Add((xIsLocal ? rawX : rawX * shape.Width, yIsLocal ? rawY : rawY * shape.Height));
                weights.Add(weight);
                suppliedKnots.Add(knot);
            }

            controlPoints.Add(end);
            weights.Add(lastWeight);
            suppliedKnots.Add(secondLastKnot);
            suppliedKnots.Add(lastKnot);
            if (controlPoints.Count <= degree) {
                degree = controlPoints.Count - 1;
            }

            if (degree < 1 || weights.Any(weight => !IsFinite(weight) || weight <= 0D)) {
                return false;
            }

            List<double> knots = TryUseSuppliedKnotVector(suppliedKnots, controlPoints.Count, degree, out List<double>? normalizedKnots)
                ? normalizedKnots!
                : BuildClampedKnotVector(controlPoints.Count, degree, firstKnot, lastKnot);
            curve = new NurbsCurve(controlPoints, weights, knots, degree);
            return true;
        }

        private static bool TryParseFunctionArguments(string? raw, string functionName, out List<string> arguments) {
            arguments = new List<string>();
            raw = NormalizeCellLiteral(raw);
            if (string.IsNullOrWhiteSpace(raw)) {
                return false;
            }

            string formula = raw!.Trim();
            if (formula.StartsWith("GUARD(", StringComparison.OrdinalIgnoreCase) && formula.EndsWith(")", StringComparison.Ordinal)) {
                formula = formula.Substring(6, formula.Length - 7).Trim();
            }

            if (!formula.StartsWith(functionName, StringComparison.OrdinalIgnoreCase)) {
                return false;
            }

            int openIndex = formula.IndexOf('(');
            int closeIndex = formula.LastIndexOf(')');
            if (openIndex < 0 || closeIndex <= openIndex) {
                return false;
            }

            string argumentsText = formula.Substring(openIndex + 1, closeIndex - openIndex - 1);
            int depth = 0;
            int start = 0;
            for (int i = 0; i < argumentsText.Length; i++) {
                char current = argumentsText[i];
                if (current == '(') {
                    depth++;
                } else if (current == ')') {
                    depth--;
                    if (depth < 0) {
                        return false;
                    }
                } else if (current == ',' && depth == 0) {
                    arguments.Add(argumentsText.Substring(start, i - start).Trim());
                    start = i + 1;
                }
            }

            arguments.Add(argumentsText.Substring(start).Trim());
            return depth == 0 && arguments.All(argument => argument.Length > 0);
        }

        private static bool TryUseSuppliedKnotVector(List<double> suppliedKnots, int controlPointCount, int degree, out List<double>? knots) {
            knots = null;
            if (suppliedKnots.Any(knot => !IsFinite(knot))) {
                return false;
            }

            for (int i = 1; i < suppliedKnots.Count; i++) {
                if (suppliedKnots[i] + 1e-9 < suppliedKnots[i - 1]) {
                    return false;
                }
            }

            int expectedFullCount = controlPointCount + degree + 1;
            if (suppliedKnots.Count == expectedFullCount) {
                knots = new List<double>(suppliedKnots);
                return IsUsableKnotDomain(knots, controlPointCount, degree);
            }

            if (suppliedKnots.Count == controlPointCount + 1) {
                List<double> expanded = new(suppliedKnots);
                double last = suppliedKnots[suppliedKnots.Count - 1];
                for (int i = 0; i < degree; i++) {
                    expanded.Add(last);
                }

                if (expanded.Count == expectedFullCount && IsUsableKnotDomain(expanded, controlPointCount, degree)) {
                    knots = expanded;
                    return true;
                }
            }

            return false;
        }

        private static bool IsUsableKnotDomain(List<double> knots, int controlPointCount, int degree) {
            if (knots.Count <= controlPointCount || knots.Count <= degree) {
                return false;
            }

            double start = knots[degree];
            double end = knots[controlPointCount];
            return IsFinite(start) && IsFinite(end) && end > start;
        }

        private static List<double> BuildClampedKnotVector(int controlPointCount, int degree, double firstKnot, double lastKnot) {
            double start = IsFinite(firstKnot) ? firstKnot : 0D;
            double end = IsFinite(lastKnot) && lastKnot > start ? lastKnot : start + 1D;
            int knotCount = controlPointCount + degree + 1;
            List<double> knots = new(knotCount);
            for (int i = 0; i < knotCount; i++) {
                if (i <= degree) {
                    knots.Add(start);
                } else if (i >= controlPointCount) {
                    knots.Add(end);
                } else {
                    double step = (i - degree) / (double)(controlPointCount - degree);
                    knots.Add(start + ((end - start) * step));
                }
            }

            return knots;
        }

        private static string? NormalizeCellLiteral(string? value) {
            if (string.IsNullOrWhiteSpace(value)) {
                return value;
            }

            string trimmed = value!.Trim();
            string normalized = trimmed.StartsWith("=", StringComparison.Ordinal)
                ? trimmed.Substring(1)
                : trimmed;
            return StripOuterGuard(normalized);
        }

        private static string StripOuterGuard(string value) {
            string normalized = value.Trim();
            while (TryGetFullFunctionArgument(normalized, "GUARD", out string? argument)) {
                normalized = argument!.Trim();
            }

            return normalized;
        }

        private static bool TryGetFullFunctionArgument(string value, string functionName, out string? argument) {
            argument = null;
            if (!value.StartsWith(functionName, StringComparison.OrdinalIgnoreCase)) {
                return false;
            }

            int openIndex = functionName.Length;
            while (openIndex < value.Length && char.IsWhiteSpace(value[openIndex])) {
                openIndex++;
            }

            if (openIndex >= value.Length || value[openIndex] != '(') {
                return false;
            }

            int depth = 0;
            for (int i = openIndex; i < value.Length; i++) {
                if (value[i] == '(') {
                    depth++;
                } else if (value[i] == ')') {
                    depth--;
                    if (depth < 0) {
                        return false;
                    }

                    if (depth == 0) {
                        if (i != value.Length - 1) {
                            return false;
                        }

                        argument = value.Substring(openIndex + 1, i - openIndex - 1);
                        return true;
                    }
                }
            }

            return false;
        }

        private static bool NearlyEqual((double X, double Y) a, (double X, double Y) b) =>
            Math.Abs(a.X - b.X) <= 1e-9 &&
            Math.Abs(a.Y - b.Y) <= 1e-9;

        private static bool IsFinite(double value) =>
            !double.IsNaN(value) &&
            !double.IsInfinity(value);

        private sealed class NurbsCurve {
            internal NurbsCurve(List<(double X, double Y)> controlPoints, List<double> weights, List<double> knots, int degree) {
                ControlPoints = controlPoints;
                Weights = weights;
                Knots = knots;
                Degree = degree;
            }

            internal List<(double X, double Y)> ControlPoints { get; }

            internal List<double> Weights { get; }

            internal List<double> Knots { get; }

            internal int Degree { get; }
        }

        private sealed class CellExpressionParser {
            private readonly string _text;
            private readonly double _width;
            private readonly double _height;
            private readonly double _locPinX;
            private readonly double _locPinY;
            private readonly double _pinX;
            private readonly double _pinY;
            private readonly double _angle;
            private int _index;

            private CellExpressionParser(string text, double width, double height, double locPinX, double locPinY, double pinX, double pinY, double angle) {
                _text = text;
                _width = width;
                _height = height;
                _locPinX = locPinX;
                _locPinY = locPinY;
                _pinX = pinX;
                _pinY = pinY;
                _angle = angle;
            }

            internal static bool TryEvaluate(string text, double width, double height, double locPinX, double locPinY, double pinX, double pinY, double angle, out double value) {
                CellExpressionParser parser = new(text, width, height, locPinX, locPinY, pinX, pinY, angle);
                if (parser.TryParseExpression(out value)) {
                    parser.SkipWhitespace();
                    if (parser._index == parser._text.Length) {
                        return true;
                    }
                }

                value = 0D;
                return false;
            }

            private bool TryParseExpression(out double value) {
                if (!TryParseTerm(out value)) {
                    return false;
                }

                while (true) {
                    SkipWhitespace();
                    if (TryRead('+')) {
                        if (!TryParseTerm(out double addend)) {
                            return false;
                        }

                        value += addend;
                    } else if (TryRead('-')) {
                        if (!TryParseTerm(out double subtrahend)) {
                            return false;
                        }

                        value -= subtrahend;
                    } else {
                        return true;
                    }
                }
            }

            private bool TryParseTerm(out double value) {
                if (!TryParsePower(out value)) {
                    return false;
                }

                while (true) {
                    SkipWhitespace();
                    if (TryRead('*')) {
                        if (!TryParsePower(out double factor)) {
                            return false;
                        }

                        value *= factor;
                    } else if (TryRead('/')) {
                        if (!TryParsePower(out double divisor) || Math.Abs(divisor) <= 1e-12) {
                            return false;
                        }

                        value /= divisor;
                    } else {
                        return true;
                    }
                }
            }

            private bool TryParsePower(out double value) {
                if (!TryParseFactor(out value)) {
                    return false;
                }

                SkipWhitespace();
                if (!TryRead('^')) {
                    return true;
                }

                if (!TryParsePower(out double exponent)) {
                    return false;
                }

                value = ShapeSheetPower(value, exponent);
                return IsFinite(value);
            }

            private bool TryParseFactor(out double value) {
                SkipWhitespace();
                if (TryRead('+')) {
                    return TryParseFactor(out value);
                }

                if (TryRead('-')) {
                    if (!TryParseFactor(out value)) {
                        return false;
                    }

                    value = -value;
                    return true;
                }

                if (TryRead('(')) {
                    if (!TryParseExpression(out value)) {
                        return false;
                    }

                    SkipWhitespace();
                    return TryRead(')');
                }

                if (TryParseIdentifier(out value)) {
                    return true;
                }

                return TryParseNumber(out value);
            }

            private bool TryParseIdentifier(out double value) {
                SkipWhitespace();
                int start = _index;
                if (_index < _text.Length && char.IsLetter(_text[_index])) {
                    _index++;
                    while (_index < _text.Length && char.IsLetterOrDigit(_text[_index])) {
                        _index++;
                    }
                }

                if (_index == start) {
                    value = 0D;
                    return false;
                }

                string identifier = _text.Substring(start, _index - start);
                SkipWhitespace();
                if (_index < _text.Length && _text[_index] == '(') {
                    return TryParseFunction(identifier, out value);
                }

                if (string.Equals(identifier, "Width", StringComparison.OrdinalIgnoreCase)) {
                    value = _width;
                    return true;
                }

                if (string.Equals(identifier, "Height", StringComparison.OrdinalIgnoreCase)) {
                    value = _height;
                    return true;
                }

                if (string.Equals(identifier, "LocPinX", StringComparison.OrdinalIgnoreCase)) {
                    value = _locPinX;
                    return true;
                }

                if (string.Equals(identifier, "LocPinY", StringComparison.OrdinalIgnoreCase)) {
                    value = _locPinY;
                    return true;
                }

                if (string.Equals(identifier, "PinX", StringComparison.OrdinalIgnoreCase)) {
                    value = _pinX;
                    return true;
                }

                if (string.Equals(identifier, "PinY", StringComparison.OrdinalIgnoreCase)) {
                    value = _pinY;
                    return true;
                }

                if (string.Equals(identifier, "Angle", StringComparison.OrdinalIgnoreCase)) {
                    value = _angle;
                    return true;
                }

                if (string.Equals(identifier, "TRUE", StringComparison.OrdinalIgnoreCase)) {
                    value = 1D;
                    return true;
                }

                if (string.Equals(identifier, "FALSE", StringComparison.OrdinalIgnoreCase)) {
                    value = 0D;
                    return true;
                }

                value = 0D;
                return false;
            }

            private bool TryParseFunction(string identifier, out double value) {
                value = 0D;
                if (string.Equals(identifier, "IF", StringComparison.OrdinalIgnoreCase)) {
                    return TryParseIfFunction(out value);
                }

                if (string.Equals(identifier, "AND", StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(identifier, "OR", StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(identifier, "NOT", StringComparison.OrdinalIgnoreCase)) {
                    return TryParseLogicalFunction(identifier, out value);
                }

                if (string.Equals(identifier, "PI", StringComparison.OrdinalIgnoreCase)) {
                    return TryParseNoArgumentFunction(Math.PI, out value);
                }

                if (string.Equals(identifier, "SIN", StringComparison.OrdinalIgnoreCase)) {
                    return TryParseSingleArgumentFunction(Math.Sin, out value);
                }

                if (string.Equals(identifier, "COS", StringComparison.OrdinalIgnoreCase)) {
                    return TryParseSingleArgumentFunction(Math.Cos, out value);
                }

                if (string.Equals(identifier, "TAN", StringComparison.OrdinalIgnoreCase)) {
                    return TryParseSingleArgumentFunction(Math.Tan, out value);
                }

                if (string.Equals(identifier, "ATAN", StringComparison.OrdinalIgnoreCase)) {
                    return TryParseSingleArgumentFunction(Math.Atan, out value);
                }

                if (string.Equals(identifier, "ATAN2", StringComparison.OrdinalIgnoreCase)) {
                    return TryParseTwoArgumentFunction((y, x) => Math.Abs(y) <= 1e-12 && Math.Abs(x) <= 1e-12 ? 0D : Math.Atan2(y, x), out value);
                }

                if (string.Equals(identifier, "RAD", StringComparison.OrdinalIgnoreCase)) {
                    return TryParseSingleArgumentFunction(angle => angle * Math.PI / 180D, out value);
                }

                if (string.Equals(identifier, "DEG", StringComparison.OrdinalIgnoreCase)) {
                    return TryParseSingleArgumentFunction(angle => angle * 180D / Math.PI, out value);
                }

                if (string.Equals(identifier, "ABS", StringComparison.OrdinalIgnoreCase)) {
                    return TryParseSingleArgumentFunction(Math.Abs, out value);
                }

                if (string.Equals(identifier, "SQRT", StringComparison.OrdinalIgnoreCase)) {
                    return TryParseSingleArgumentFunction(Math.Sqrt, out value);
                }

                if (string.Equals(identifier, "INT", StringComparison.OrdinalIgnoreCase)) {
                    return TryParseSingleArgumentFunction(Math.Floor, out value);
                }

                if (string.Equals(identifier, "POW", StringComparison.OrdinalIgnoreCase)) {
                    return TryParseTwoArgumentFunction(ShapeSheetPower, out value);
                }

                if (string.Equals(identifier, "ROUND", StringComparison.OrdinalIgnoreCase)) {
                    return TryParseTwoArgumentFunction(RoundShapeSheetValue, out value);
                }

                bool isMin = string.Equals(identifier, "MIN", StringComparison.OrdinalIgnoreCase);
                bool isMax = string.Equals(identifier, "MAX", StringComparison.OrdinalIgnoreCase);
                if (!isMin && !isMax) {
                    return false;
                }

                if (!TryRead('(')) {
                    return false;
                }

                bool hasValue = false;
                double result = isMin
                    ? double.PositiveInfinity
                    : double.NegativeInfinity;
                while (true) {
                    if (!TryParseExpression(out double argument)) {
                        return false;
                    }

                    hasValue = true;
                    result = isMin
                        ? Math.Min(result, argument)
                        : Math.Max(result, argument);
                    SkipWhitespace();
                    if (TryRead(')')) {
                        value = result;
                        return hasValue && IsFinite(value);
                    }

                    if (!TryRead(',')) {
                        return false;
                    }
                }
            }

            private bool TryParseNoArgumentFunction(double result, out double value) {
                value = 0D;
                if (!TryRead('(')) {
                    return false;
                }

                SkipWhitespace();
                if (!TryRead(')')) {
                    return false;
                }

                value = result;
                return IsFinite(value);
            }

            private bool TryParseSingleArgumentFunction(Func<double, double> function, out double value) {
                value = 0D;
                if (!TryRead('(') ||
                    !TryParseExpression(out double argument)) {
                    return false;
                }

                SkipWhitespace();
                if (!TryRead(')')) {
                    return false;
                }

                value = function(argument);
                return IsFinite(value);
            }

            private bool TryParseTwoArgumentFunction(Func<double, double, double> function, out double value) {
                value = 0D;
                if (!TryRead('(') ||
                    !TryParseExpression(out double firstArgument) ||
                    !TryReadComma() ||
                    !TryParseExpression(out double secondArgument)) {
                    return false;
                }

                SkipWhitespace();
                if (!TryRead(')')) {
                    return false;
                }

                value = function(firstArgument, secondArgument);
                return IsFinite(value);
            }

            private static double ShapeSheetPower(double number, double exponent) {
                if (Math.Abs(number) <= 1e-12 && exponent <= 0D) {
                    return 0D;
                }

                if (number < 0D && !IsNearlyInteger(exponent)) {
                    return 0D;
                }

                return Math.Pow(number, exponent);
            }

            private static double RoundShapeSheetValue(double number, double numberOfDigits) {
                int digits = (int)Math.Round(numberOfDigits, MidpointRounding.AwayFromZero);
                double factor = Math.Pow(10D, Math.Abs(digits));
                if (!IsFinite(factor) || Math.Abs(factor) <= 1e-12) {
                    return double.NaN;
                }

                if (digits >= 0) {
                    return Math.Round(number * factor, 0, MidpointRounding.AwayFromZero) / factor;
                }

                return Math.Round(number / factor, 0, MidpointRounding.AwayFromZero) * factor;
            }

            private static bool IsNearlyInteger(double value) =>
                Math.Abs(value - Math.Round(value)) <= 1e-9;

            private bool TryParseLogicalFunction(string identifier, out double value) {
                value = 0D;
                if (!TryReadFunctionArguments(out List<string> arguments)) {
                    return false;
                }

                bool result;
                if (string.Equals(identifier, "AND", StringComparison.OrdinalIgnoreCase)) {
                    if (arguments.Count == 0) {
                        return false;
                    }

                    result = true;
                    foreach (string argument in arguments) {
                        if (!TryEvaluateCondition(argument, _width, _height, _locPinX, _locPinY, _pinX, _pinY, _angle, out bool argumentValue)) {
                            return false;
                        }

                        result &= argumentValue;
                    }
                } else if (string.Equals(identifier, "OR", StringComparison.OrdinalIgnoreCase)) {
                    if (arguments.Count == 0) {
                        return false;
                    }

                    result = false;
                    foreach (string argument in arguments) {
                        if (!TryEvaluateCondition(argument, _width, _height, _locPinX, _locPinY, _pinX, _pinY, _angle, out bool argumentValue)) {
                            return false;
                        }

                        result |= argumentValue;
                    }
                } else if (string.Equals(identifier, "NOT", StringComparison.OrdinalIgnoreCase)) {
                    if (arguments.Count != 1 ||
                        !TryEvaluateCondition(arguments[0], _width, _height, _locPinX, _locPinY, _pinX, _pinY, _angle, out bool argumentValue)) {
                        return false;
                    }

                    result = !argumentValue;
                } else {
                    return false;
                }

                value = result ? 1D : 0D;
                return true;
            }

            private bool TryParseIfFunction(out double value) {
                value = 0D;
                if (!TryRead('(') ||
                    !TryParseCondition(out bool condition) ||
                    !TryReadComma() ||
                    !TryReadFunctionArgument(',', out string? whenTrueExpression) ||
                    !TryReadFunctionArgument(')', out string? whenFalseExpression)) {
                    return false;
                }

                string selectedExpression = condition ? whenTrueExpression! : whenFalseExpression!;
                return TryEvaluate(selectedExpression, _width, _height, _locPinX, _locPinY, _pinX, _pinY, _angle, out value) &&
                       IsFinite(value);
            }

            private static bool TryEvaluateCondition(string text, double width, double height, double locPinX, double locPinY, double pinX, double pinY, double angle, out bool value) {
                CellExpressionParser parser = new(text, width, height, locPinX, locPinY, pinX, pinY, angle);
                if (parser.TryParseCondition(out value)) {
                    parser.SkipWhitespace();
                    if (parser._index == parser._text.Length) {
                        return true;
                    }
                }

                value = false;
                return false;
            }

            private bool TryParseCondition(out bool value) {
                value = false;
                if (!TryParseExpression(out double left)) {
                    return false;
                }

                SkipWhitespace();
                if (!TryReadComparisonOperator(out string? comparisonOperator)) {
                    value = Math.Abs(left) > 1e-9;
                    return true;
                }

                if (!TryParseExpression(out double right)) {
                    return false;
                }

                switch (comparisonOperator) {
                    case "<":
                        value = left < right;
                        return true;
                    case "<=":
                        value = left <= right;
                        return true;
                    case ">":
                        value = left > right;
                        return true;
                    case ">=":
                        value = left >= right;
                        return true;
                    case "=":
                        value = Math.Abs(left - right) <= 1e-9;
                        return true;
                    case "<>":
                    case "!=":
                        value = Math.Abs(left - right) > 1e-9;
                        return true;
                    default:
                        return false;
                }
            }

            private bool TryReadComparisonOperator(out string? comparisonOperator) {
                comparisonOperator = null;
                if (_index >= _text.Length) {
                    return false;
                }

                if (_index + 1 < _text.Length) {
                    string pair = _text.Substring(_index, 2);
                    if (pair == "<=" || pair == ">=" || pair == "<>" || pair == "!=") {
                        comparisonOperator = pair;
                        _index += 2;
                        return true;
                    }
                }

                char current = _text[_index];
                if (current == '<' || current == '>' || current == '=') {
                    comparisonOperator = current.ToString();
                    _index++;
                    return true;
                }

                return false;
            }

            private bool TryReadComma() {
                SkipWhitespace();
                return TryRead(',');
            }

            private bool TryReadFunctionArguments(out List<string> arguments) {
                arguments = new List<string>();
                if (!TryRead('(')) {
                    return false;
                }

                int start = _index;
                int depth = 0;
                while (_index < _text.Length) {
                    char current = _text[_index];
                    if (current == '(') {
                        depth++;
                    } else if (current == ')') {
                        if (depth == 0) {
                            string argument = _text.Substring(start, _index - start).Trim();
                            if (argument.Length == 0) {
                                return false;
                            }

                            arguments.Add(argument);
                            _index++;
                            return true;
                        }

                        depth--;
                    } else if (current == ',' && depth == 0) {
                        string argument = _text.Substring(start, _index - start).Trim();
                        if (argument.Length == 0) {
                            return false;
                        }

                        arguments.Add(argument);
                        _index++;
                        start = _index;
                        continue;
                    }

                    _index++;
                }

                return false;
            }

            private bool TryReadFunctionArgument(char delimiter, out string? argument) {
                SkipWhitespace();
                int start = _index;
                int depth = 0;
                while (_index < _text.Length) {
                    char current = _text[_index];
                    if (current == '(') {
                        depth++;
                    } else if (current == ')') {
                        if (depth == 0) {
                            if (delimiter != ')') {
                                argument = null;
                                return false;
                            }

                            argument = _text.Substring(start, _index - start).Trim();
                            _index++;
                            return argument.Length > 0;
                        }

                        depth--;
                    } else if (current == ',' && depth == 0) {
                        if (delimiter != ',') {
                            argument = null;
                            return false;
                        }

                        argument = _text.Substring(start, _index - start).Trim();
                        _index++;
                        return argument.Length > 0;
                    }

                    _index++;
                }

                argument = null;
                return false;
            }

            private bool TryParseNumber(out double value) {
                SkipWhitespace();
                int start = _index;
                bool hasDigit = false;
                while (_index < _text.Length && char.IsDigit(_text[_index])) {
                    _index++;
                    hasDigit = true;
                }

                if (_index < _text.Length && _text[_index] == '.') {
                    _index++;
                    while (_index < _text.Length && char.IsDigit(_text[_index])) {
                        _index++;
                        hasDigit = true;
                    }
                }

                if (!hasDigit) {
                    value = 0D;
                    return false;
                }

                if (_index < _text.Length && (_text[_index] == 'e' || _text[_index] == 'E')) {
                    int exponentStart = _index;
                    _index++;
                    if (_index < _text.Length && (_text[_index] == '+' || _text[_index] == '-')) {
                        _index++;
                    }

                    int exponentDigitsStart = _index;
                    while (_index < _text.Length && char.IsDigit(_text[_index])) {
                        _index++;
                    }

                    if (_index == exponentDigitsStart) {
                        _index = exponentStart;
                    }
                }

                if (!double.TryParse(_text.Substring(start, _index - start), NumberStyles.Float, CultureInfo.InvariantCulture, out value)) {
                    return false;
                }

                TryApplyUnitSuffix(ref value);
                return IsFinite(value);
            }

            private void TryApplyUnitSuffix(ref double value) {
                int suffixStart = _index;
                SkipWhitespace();
                if (_index < _text.Length && _text[_index] == '%') {
                    value *= 0.01D;
                    _index++;
                    return;
                }

                int unitStart = _index;
                while (_index < _text.Length && char.IsLetter(_text[_index])) {
                    _index++;
                }

                if (_index == unitStart) {
                    _index = suffixStart;
                    return;
                }

                string unit = _text.Substring(unitStart, _index - unitStart);
                if (TryGetUnitScale(unit, out double scale)) {
                    value *= scale;
                    return;
                }

                _index = suffixStart;
            }

            private static bool TryGetUnitScale(string unit, out double scale) {
                switch (unit.ToLowerInvariant()) {
                    case "deg":
                    case "degree":
                    case "degrees":
                        scale = Math.PI / 180D;
                        return true;
                    case "rad":
                    case "radian":
                    case "radians":
                    case "in":
                    case "inch":
                    case "inches":
                        scale = 1D;
                        return true;
                    case "ft":
                    case "foot":
                    case "feet":
                        scale = 12D;
                        return true;
                    case "yd":
                    case "yard":
                    case "yards":
                        scale = 36D;
                        return true;
                    case "mi":
                    case "mile":
                    case "miles":
                        scale = 63360D;
                        return true;
                    case "mm":
                        scale = 1D / 25.4D;
                        return true;
                    case "cm":
                        scale = 1D / 2.54D;
                        return true;
                    case "m":
                    case "meter":
                    case "meters":
                        scale = 100D / 2.54D;
                        return true;
                    case "km":
                    case "kilometer":
                    case "kilometers":
                        scale = 100000D / 2.54D;
                        return true;
                    case "pt":
                    case "point":
                    case "points":
                        scale = 1D / 72D;
                        return true;
                    case "pc":
                    case "pica":
                    case "picas":
                        scale = 1D / 6D;
                        return true;
                    default:
                        scale = 1D;
                        return false;
                }
            }

            private bool TryRead(char expected) {
                if (_index < _text.Length && _text[_index] == expected) {
                    _index++;
                    return true;
                }

                return false;
            }

            private void SkipWhitespace() {
                while (_index < _text.Length && char.IsWhiteSpace(_text[_index])) {
                    _index++;
                }
            }
        }
    }
}
