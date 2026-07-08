using System;
using System.Collections.Generic;
using System.Globalization;
using OfficeIMO.Drawing;

namespace OfficeIMO.Visio {
    internal static partial class VisioSvgPreviewRasterizer {
        private static bool TryParsePath(string? data, out List<SvgPathContour> contours) {
            contours = new List<SvgPathContour>();
            if (string.IsNullOrWhiteSpace(data)) {
                return false;
            }

            List<PathToken> tokens = TokenizePath(data!);
            if (tokens.Count == 0) {
                return false;
            }

            List<(double X, double Y)> currentPath = new();
            double x = 0D;
            double y = 0D;
            double startX = 0D;
            double startY = 0D;
            (double X, double Y)? lastCubicControl = null;
            (double X, double Y)? lastQuadraticControl = null;
            char command = '\0';
            int index = 0;
            while (index < tokens.Count) {
                if (tokens[index].Command != '\0') {
                    command = tokens[index].Command;
                    index++;
                }

                bool relative = char.IsLower(command);
                char upper = char.ToUpperInvariant(command);
                if (upper == 'Z') {
                    if (currentPath.Count > 0) {
                        currentPath.Add((startX, startY));
                        contours.Add(new SvgPathContour(currentPath, isClosed: true));
                        currentPath = new List<(double X, double Y)>();
                    }

                    x = startX;
                    y = startY;
                    lastCubicControl = null;
                    lastQuadraticControl = null;
                    continue;
                }

                if (upper == 'M') {
                    if (!TryReadPathNumber(tokens, ref index, out double nextX) ||
                        !TryReadPathNumber(tokens, ref index, out double nextY)) {
                        return false;
                    }

                    AddOpenContour(contours, currentPath);
                    x = relative ? x + nextX : nextX;
                    y = relative ? y + nextY : nextY;
                    startX = x;
                    startY = y;
                    currentPath = new List<(double X, double Y)> { (x, y) };
                    command = relative ? 'l' : 'L';
                    lastCubicControl = null;
                    lastQuadraticControl = null;
                    continue;
                }

                if (upper == 'L') {
                    if (!TryReadPathNumber(tokens, ref index, out double nextX) ||
                        !TryReadPathNumber(tokens, ref index, out double nextY)) {
                        return false;
                    }

                    x = relative ? x + nextX : nextX;
                    y = relative ? y + nextY : nextY;
                    currentPath.Add((x, y));
                    lastCubicControl = null;
                    lastQuadraticControl = null;
                    continue;
                }

                if (upper == 'H') {
                    if (!TryReadPathNumber(tokens, ref index, out double nextX)) {
                        return false;
                    }

                    x = relative ? x + nextX : nextX;
                    currentPath.Add((x, y));
                    lastCubicControl = null;
                    lastQuadraticControl = null;
                    continue;
                }

                if (upper == 'V') {
                    if (!TryReadPathNumber(tokens, ref index, out double nextY)) {
                        return false;
                    }

                    y = relative ? y + nextY : nextY;
                    currentPath.Add((x, y));
                    lastCubicControl = null;
                    lastQuadraticControl = null;
                    continue;
                }

                if (upper == 'C') {
                    if (!TryReadPathNumber(tokens, ref index, out double c1x) ||
                        !TryReadPathNumber(tokens, ref index, out double c1y) ||
                        !TryReadPathNumber(tokens, ref index, out double c2x) ||
                        !TryReadPathNumber(tokens, ref index, out double c2y) ||
                        !TryReadPathNumber(tokens, ref index, out double endX) ||
                        !TryReadPathNumber(tokens, ref index, out double endY)) {
                        return false;
                    }

                    (double X, double Y) p1 = (relative ? x + c1x : c1x, relative ? y + c1y : c1y);
                    (double X, double Y) p2 = (relative ? x + c2x : c2x, relative ? y + c2y : c2y);
                    (double X, double Y) end = (relative ? x + endX : endX, relative ? y + endY : endY);
                    AppendCubic(currentPath, (x, y), p1, p2, end);
                    x = end.X;
                    y = end.Y;
                    lastCubicControl = p2;
                    lastQuadraticControl = null;
                    continue;
                }

                if (upper == 'S') {
                    if (!TryReadPathNumber(tokens, ref index, out double c2x) ||
                        !TryReadPathNumber(tokens, ref index, out double c2y) ||
                        !TryReadPathNumber(tokens, ref index, out double endX) ||
                        !TryReadPathNumber(tokens, ref index, out double endY)) {
                        return false;
                    }

                    (double X, double Y) p1 = lastCubicControl.HasValue
                        ? ((2D * x) - lastCubicControl.Value.X, (2D * y) - lastCubicControl.Value.Y)
                        : (x, y);
                    (double X, double Y) p2 = (relative ? x + c2x : c2x, relative ? y + c2y : c2y);
                    (double X, double Y) end = (relative ? x + endX : endX, relative ? y + endY : endY);
                    AppendCubic(currentPath, (x, y), p1, p2, end);
                    x = end.X;
                    y = end.Y;
                    lastCubicControl = p2;
                    lastQuadraticControl = null;
                    continue;
                }

                if (upper == 'Q') {
                    if (!TryReadPathNumber(tokens, ref index, out double cx) ||
                        !TryReadPathNumber(tokens, ref index, out double cy) ||
                        !TryReadPathNumber(tokens, ref index, out double endX) ||
                        !TryReadPathNumber(tokens, ref index, out double endY)) {
                        return false;
                    }

                    (double X, double Y) control = (relative ? x + cx : cx, relative ? y + cy : cy);
                    (double X, double Y) end = (relative ? x + endX : endX, relative ? y + endY : endY);
                    AppendQuadratic(currentPath, (x, y), control, end);
                    x = end.X;
                    y = end.Y;
                    lastCubicControl = null;
                    lastQuadraticControl = control;
                    continue;
                }

                if (upper == 'T') {
                    if (!TryReadPathNumber(tokens, ref index, out double endX) ||
                        !TryReadPathNumber(tokens, ref index, out double endY)) {
                        return false;
                    }

                    (double X, double Y) control = lastQuadraticControl.HasValue
                        ? ((2D * x) - lastQuadraticControl.Value.X, (2D * y) - lastQuadraticControl.Value.Y)
                        : (x, y);
                    (double X, double Y) end = (relative ? x + endX : endX, relative ? y + endY : endY);
                    AppendQuadratic(currentPath, (x, y), control, end);
                    x = end.X;
                    y = end.Y;
                    lastCubicControl = null;
                    lastQuadraticControl = control;
                    continue;
                }

                if (upper == 'A') {
                    if (!TryReadPathNumber(tokens, ref index, out double rx) ||
                        !TryReadPathNumber(tokens, ref index, out double ry) ||
                        !TryReadPathNumber(tokens, ref index, out double xAxisRotation) ||
                        !TryReadPathNumber(tokens, ref index, out double largeArcFlag) ||
                        !TryReadPathNumber(tokens, ref index, out double sweepFlag) ||
                        !TryReadPathNumber(tokens, ref index, out double endX) ||
                        !TryReadPathNumber(tokens, ref index, out double endY)) {
                        return false;
                    }

                    (double X, double Y) end = (relative ? x + endX : endX, relative ? y + endY : endY);
                    AppendSvgArc(currentPath, (x, y), end, rx, ry, xAxisRotation, Math.Abs(largeArcFlag) > 0.5D, Math.Abs(sweepFlag) > 0.5D);
                    x = end.X;
                    y = end.Y;
                    lastCubicControl = null;
                    lastQuadraticControl = null;
                    continue;
                }

                return false;
            }

            AddOpenContour(contours, currentPath);
            return contours.Count > 0;
        }

        private static void AddOpenContour(List<SvgPathContour> contours, List<(double X, double Y)> points) {
            if (points.Count > 0) {
                contours.Add(new SvgPathContour(points, isClosed: false));
            }
        }

        private static void AppendCubic(List<(double X, double Y)> points, (double X, double Y) start, (double X, double Y) c1, (double X, double Y) c2, (double X, double Y) end) {
            for (int i = 1; i <= 18; i++) {
                double t = i / 18D;
                double u = 1D - t;
                points.Add((
                    (u * u * u * start.X) + (3D * u * u * t * c1.X) + (3D * u * t * t * c2.X) + (t * t * t * end.X),
                    (u * u * u * start.Y) + (3D * u * u * t * c1.Y) + (3D * u * t * t * c2.Y) + (t * t * t * end.Y)));
            }
        }

        private static void AppendQuadratic(List<(double X, double Y)> points, (double X, double Y) start, (double X, double Y) control, (double X, double Y) end) {
            for (int i = 1; i <= 14; i++) {
                double t = i / 14D;
                double u = 1D - t;
                points.Add((
                    (u * u * start.X) + (2D * u * t * control.X) + (t * t * end.X),
                    (u * u * start.Y) + (2D * u * t * control.Y) + (t * t * end.Y)));
            }
        }

        private static void AppendSvgArc(
            List<(double X, double Y)> points,
            (double X, double Y) start,
            (double X, double Y) end,
            double radiusX,
            double radiusY,
            double rotationDegrees,
            bool largeArc,
            bool sweep) {
            radiusX = Math.Abs(radiusX);
            radiusY = Math.Abs(radiusY);
            if (radiusX <= 0D || radiusY <= 0D || NearlyEqual(start, end)) {
                points.Add(end);
                return;
            }

            double rotationRadians = OfficeGeometry.DegreesToRadians(rotationDegrees);
            double cosPhi = Math.Cos(rotationRadians);
            double sinPhi = Math.Sin(rotationRadians);
            double dx2 = (start.X - end.X) / 2D;
            double dy2 = (start.Y - end.Y) / 2D;
            double x1Prime = (cosPhi * dx2) + (sinPhi * dy2);
            double y1Prime = (-sinPhi * dx2) + (cosPhi * dy2);

            double radiiScale = ((x1Prime * x1Prime) / (radiusX * radiusX)) + ((y1Prime * y1Prime) / (radiusY * radiusY));
            if (radiiScale > 1D) {
                double scale = Math.Sqrt(radiiScale);
                radiusX *= scale;
                radiusY *= scale;
            }

            double rxSquared = radiusX * radiusX;
            double rySquared = radiusY * radiusY;
            double x1Squared = x1Prime * x1Prime;
            double y1Squared = y1Prime * y1Prime;
            double denominator = (rxSquared * y1Squared) + (rySquared * x1Squared);
            if (denominator <= 0D) {
                points.Add(end);
                return;
            }

            double numerator = Math.Max(0D, (rxSquared * rySquared) - (rxSquared * y1Squared) - (rySquared * x1Squared));
            double factor = (largeArc == sweep ? -1D : 1D) * Math.Sqrt(numerator / denominator);
            double centerXPrime = factor * radiusX * y1Prime / radiusY;
            double centerYPrime = -factor * radiusY * x1Prime / radiusX;
            double centerX = (cosPhi * centerXPrime) - (sinPhi * centerYPrime) + ((start.X + end.X) / 2D);
            double centerY = (sinPhi * centerXPrime) + (cosPhi * centerYPrime) + ((start.Y + end.Y) / 2D);

            double startAngle = VectorAngle(1D, 0D, (x1Prime - centerXPrime) / radiusX, (y1Prime - centerYPrime) / radiusY);
            double sweepAngle = VectorAngle(
                (x1Prime - centerXPrime) / radiusX,
                (y1Prime - centerYPrime) / radiusY,
                (-x1Prime - centerXPrime) / radiusX,
                (-y1Prime - centerYPrime) / radiusY);

            if (!sweep && sweepAngle > 0D) {
                sweepAngle -= Math.PI * 2D;
            } else if (sweep && sweepAngle < 0D) {
                sweepAngle += Math.PI * 2D;
            }

            int segments = Math.Max(4, (int)Math.Ceiling(Math.Abs(sweepAngle) / (Math.PI / 12D)));
            points.AddRange(OfficeGeometry.CreateEllipticalArcPointsAsTuples(
                centerX,
                centerY,
                radiusX,
                radiusY,
                startAngle,
                sweepAngle,
                segments,
                rotationRadians,
                centerX,
                centerY));
        }

        private static double VectorAngle(double ux, double uy, double vx, double vy) {
            double dot = (ux * vx) + (uy * vy);
            double length = Math.Sqrt(((ux * ux) + (uy * uy)) * ((vx * vx) + (vy * vy)));
            if (length <= 0D) {
                return 0D;
            }

            double ratio = Math.Max(-1D, Math.Min(1D, dot / length));
            double angle = Math.Acos(ratio);
            return ((ux * vy) - (uy * vx)) < 0D ? -angle : angle;
        }

        private static bool NearlyEqual((double X, double Y) first, (double X, double Y) second) =>
            Math.Abs(first.X - second.X) < 0.000001D && Math.Abs(first.Y - second.Y) < 0.000001D;

        private static bool TryReadPathNumber(IReadOnlyList<PathToken> tokens, ref int index, out double value) {
            value = 0D;
            if (index >= tokens.Count || tokens[index].Command != '\0') {
                return false;
            }

            value = tokens[index].Number;
            index++;
            return true;
        }

        private static List<PathToken> TokenizePath(string value) {
            List<PathToken> tokens = new();
            int index = 0;
            while (index < value.Length) {
                char current = value[index];
                if (char.IsWhiteSpace(current) || current == ',') {
                    index++;
                    continue;
                }

                if (char.IsLetter(current)) {
                    tokens.Add(new PathToken(current, 0D));
                    index++;
                    continue;
                }

                int start = index;
                index++;
                while (index < value.Length && (char.IsDigit(value[index]) || value[index] == '.' || value[index] == 'e' || value[index] == 'E' || value[index] == '-' || value[index] == '+')) {
                    if ((value[index] == '-' || value[index] == '+') && value[index - 1] != 'e' && value[index - 1] != 'E') {
                        break;
                    }

                    index++;
                }

                if (double.TryParse(value.Substring(start, index - start), NumberStyles.Float, CultureInfo.InvariantCulture, out double number)) {
                    tokens.Add(new PathToken('\0', number));
                }
            }

            return tokens;
        }

        private readonly struct SvgPathContour {
            internal SvgPathContour(List<(double X, double Y)> points, bool isClosed) {
                Points = points;
                IsClosed = isClosed;
            }

            internal List<(double X, double Y)> Points { get; }

            internal bool IsClosed { get; }
        }

        private readonly struct PathToken {
            internal PathToken(char command, double number) {
                Command = command;
                Number = number;
            }

            internal char Command { get; }

            internal double Number { get; }
        }
    }
}
