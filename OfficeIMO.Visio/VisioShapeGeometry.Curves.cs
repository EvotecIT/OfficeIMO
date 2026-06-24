using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Xml.Linq;
using OfficeIMO.Drawing;

namespace OfficeIMO.Visio {
    internal static partial class VisioShapeGeometry {

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

            path.AddRange(OfficeGeometry.CreateEllipticalArcPointsAsTuples(
                centerX,
                centerY,
                radius,
                radius,
                startAngle,
                sweep,
                ArcSegmentCount));
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
            foreach ((double X, double Y) point in OfficeGeometry.CreateEllipticalArcPointsAsTuples(
                center.X,
                center.Y,
                radius,
                radius,
                startAngle,
                sweep,
                ArcSegmentCount)) {
                path.Add(InverseTransformEllipsePoint(point, inverseCos, inverseSin, ratio));
            }
        }

        private static void AppendCubicBezier(
            List<(double X, double Y)> path,
            (double X, double Y) start,
            (double X, double Y) control1,
            (double X, double Y) control2,
            (double X, double Y) end) {
            path.AddRange(OfficeGeometry.CreateCubicBezierPoints(start, control1, control2, end, ArcSegmentCount));
        }

        private static void AppendQuadraticBezier(
            List<(double X, double Y)> path,
            (double X, double Y) start,
            (double X, double Y) control,
            (double X, double Y) end) {
            path.AddRange(OfficeGeometry.CreateQuadraticBezierPoints(start, control, end, ArcSegmentCount));
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

    }
}
