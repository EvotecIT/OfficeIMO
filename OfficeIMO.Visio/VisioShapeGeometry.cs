using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Xml.Linq;
using OfficeIMO.Drawing;

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


    internal static partial class VisioShapeGeometry {
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
                    paths.Clear();
                    return false;
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
            points.AddRange(OfficeGeometry.CreateEllipticalArcPointsAsTuples(
                centerX,
                centerY,
                radius,
                radius,
                startAngle,
                endAngle - startAngle,
                ArcSegmentCount));
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
    }
}
