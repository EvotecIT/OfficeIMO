using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using DocumentFormat.OpenXml.Presentation;
using OfficeIMO.Drawing;
using A = DocumentFormat.OpenXml.Drawing;
using static OfficeIMO.PowerPoint.PowerPointDrawingValueValidator;

namespace OfficeIMO.PowerPoint {
    public partial class PowerPointSlide {
        private const long CustomGeometryCoordinateSize = 100000L;
        private const int MaximumCustomGeometryCommands = 20000;
        private const int MaximumEvenOddIntersectionSegments = 2048;

        /// <summary>
        /// Adds an editable PowerPoint custom-geometry shape from a shared drawing path or polygon.
        /// The shared descriptor supplies the geometry and basic fill/stroke styling; the supplied
        /// bounds control its size and placement on the slide.
        /// </summary>
        public PowerPointAutoShape AddCustomGeometry(OfficeShape geometry, long left, long top,
            long width, long height, string? name = null) {
            if (geometry == null) throw new ArgumentNullException(nameof(geometry));
            if (width <= 0) throw new ArgumentOutOfRangeException(nameof(width));
            if (height <= 0) throw new ArgumentOutOfRangeException(nameof(height));
            ValidateCustomGeometryTransformCoordinate(left, nameof(left));
            ValidateCustomGeometryTransformCoordinate(top, nameof(top));
            ValidateCustomGeometryExtent(width, nameof(width));
            ValidateCustomGeometryExtent(height, nameof(height));
            ValidateCustomGeometryOpacity(geometry.FillOpacity,
                nameof(geometry.FillOpacity));
            ValidateCustomGeometryOpacity(geometry.StrokeOpacity,
                nameof(geometry.StrokeOpacity));

            IReadOnlyList<OfficePathCommand> commands = GetCustomGeometryCommands(geometry);
            if (commands.Count > MaximumCustomGeometryCommands) {
                throw new ArgumentException(
                    $"Custom geometry supports at most {MaximumCustomGeometryCommands} path commands.",
                    nameof(geometry));
            }
            if (geometry.FillColor.HasValue
                && geometry.FillRule == OfficeFillRule.EvenOdd
                && !CanEncodeEvenOddFillAsNonZero(commands)) {
                throw new NotSupportedException(
                    "DrawingML custom geometry cannot faithfully encode the even-odd fill rule. Use a non-zero fill rule, remove the fill, or split the geometry into separate shapes.");
            }
            if (geometry.FillGradient != null
                || geometry.FillRadialGradient != null
                || geometry.StrokeGradient != null
                || geometry.StrokeRadialGradient != null) {
                throw new NotSupportedException(
                    "PowerPoint custom geometry does not yet project shared fill or stroke gradients. Remove the gradients or use a supported solid style.");
            }
            if (geometry.StrokeColor.HasValue) {
                PowerPointShape.ToDrawingLineWidth(geometry.StrokeWidth,
                    nameof(geometry));
            }
            A.CustomGeometry customGeometry = CreateCustomGeometry(
                geometry, commands);

            PowerPointAutoShape result = AddShape(
                OfficePresetShapeType.Rectangle, left, top, width, height, name);
            Shape shape = (Shape)result.Element;
            ShapeProperties properties = shape.ShapeProperties
                ?? throw new InvalidOperationException("The added shape has no shape properties.");
            A.Transform2D transform = properties.GetFirstChild<A.Transform2D>()
                ?? throw new InvalidOperationException("The added shape has no transform.");
            properties.RemoveAllChildren<A.PresetGeometry>();
            properties.InsertAfter(customGeometry, transform);
            ApplyCustomGeometryStyle(result, properties, geometry);
            return result;
        }

        private static bool CanEncodeEvenOddFillAsNonZero(
            IReadOnlyList<OfficePathCommand> commands) {
            if (commands.Count(command => command.Kind
                    == OfficePathCommandKind.MoveTo) != 1
                || commands.Any(command => command.Kind is not (
                    OfficePathCommandKind.MoveTo or OfficePathCommandKind.LineTo
                    or OfficePathCommandKind.Close))) {
                return false;
            }

            List<OfficePoint> points = commands
                .Where(command => command.Kind is OfficePathCommandKind.MoveTo
                    or OfficePathCommandKind.LineTo)
                .Select(command => command.Point)
                .ToList();
            if (points.Count > 1 && PointsEqual(points[0],
                    points[points.Count - 1])) {
                points.RemoveAt(points.Count - 1);
            }
            if (points.Count < 3 || points.Select(point => (point.X, point.Y))
                    .Distinct().Count() != points.Count) {
                return false;
            }
            if (points.Count > MaximumEvenOddIntersectionSegments) {
                return false;
            }

            double signedArea = 0D;
            for (int index = 0; index < points.Count; index++) {
                OfficePoint current = points[index];
                OfficePoint next = points[(index + 1) % points.Count];
                signedArea += current.X * next.Y - next.X * current.Y;
            }
            if (Math.Abs(signedArea) <= 1E-9D) return false;

            for (int first = 0; first < points.Count; first++) {
                int firstNext = (first + 1) % points.Count;
                for (int second = first + 1; second < points.Count; second++) {
                    int secondNext = (second + 1) % points.Count;
                    if (first == second || firstNext == second
                        || secondNext == first) continue;
                    if (SegmentsIntersect(points[first], points[firstNext],
                            points[second], points[secondNext])) return false;
                }
            }
            return true;
        }

        private static bool SegmentsIntersect(OfficePoint firstStart,
            OfficePoint firstEnd, OfficePoint secondStart,
            OfficePoint secondEnd) {
            double firstSide = Cross(firstStart, firstEnd, secondStart);
            double secondSide = Cross(firstStart, firstEnd, secondEnd);
            double thirdSide = Cross(secondStart, secondEnd, firstStart);
            double fourthSide = Cross(secondStart, secondEnd, firstEnd);
            const double epsilon = 1E-9D;
            if (Math.Abs(firstSide) <= epsilon && IsOnSegment(firstStart,
                    firstEnd, secondStart, epsilon)) return true;
            if (Math.Abs(secondSide) <= epsilon && IsOnSegment(firstStart,
                    firstEnd, secondEnd, epsilon)) return true;
            if (Math.Abs(thirdSide) <= epsilon && IsOnSegment(secondStart,
                    secondEnd, firstStart, epsilon)) return true;
            if (Math.Abs(fourthSide) <= epsilon && IsOnSegment(secondStart,
                    secondEnd, firstEnd, epsilon)) return true;
            return (firstSide > 0D) != (secondSide > 0D)
                && (thirdSide > 0D) != (fourthSide > 0D);
        }

        private static bool IsOnSegment(OfficePoint start, OfficePoint end,
            OfficePoint point, double epsilon) => point.X
                >= Math.Min(start.X, end.X) - epsilon
                && point.X <= Math.Max(start.X, end.X) + epsilon
                && point.Y >= Math.Min(start.Y, end.Y) - epsilon
                && point.Y <= Math.Max(start.Y, end.Y) + epsilon;

        private static double Cross(OfficePoint start, OfficePoint end,
            OfficePoint point) => (end.X - start.X) * (point.Y - start.Y)
                - (end.Y - start.Y) * (point.X - start.X);

        private static bool PointsEqual(OfficePoint first,
            OfficePoint second) => first.X.Equals(second.X)
                && first.Y.Equals(second.Y);

        /// <summary>Adds shared custom geometry using point measurements.</summary>
        public PowerPointAutoShape AddCustomGeometryPoints(OfficeShape geometry, double leftPoints,
            double topPoints, double widthPoints, double heightPoints, string? name = null) {
            return AddCustomGeometry(geometry,
                PowerPointUnits.FromPoints(leftPoints),
                PowerPointUnits.FromPoints(topPoints),
                PowerPointUnits.FromPoints(widthPoints),
                PowerPointUnits.FromPoints(heightPoints),
                name);
        }

        /// <summary>Adds shared custom geometry using centimeter measurements.</summary>
        public PowerPointAutoShape AddCustomGeometryCm(OfficeShape geometry, double leftCm,
            double topCm, double widthCm, double heightCm, string? name = null) {
            return AddCustomGeometry(geometry,
                PowerPointUnits.FromCentimeters(leftCm),
                PowerPointUnits.FromCentimeters(topCm),
                PowerPointUnits.FromCentimeters(widthCm),
                PowerPointUnits.FromCentimeters(heightCm),
                name);
        }

        /// <summary>Adds shared custom geometry using inch measurements.</summary>
        public PowerPointAutoShape AddCustomGeometryInches(OfficeShape geometry, double leftInches,
            double topInches, double widthInches, double heightInches, string? name = null) {
            return AddCustomGeometry(geometry,
                PowerPointUnits.FromInches(leftInches),
                PowerPointUnits.FromInches(topInches),
                PowerPointUnits.FromInches(widthInches),
                PowerPointUnits.FromInches(heightInches),
                name);
        }

        private static IReadOnlyList<OfficePathCommand> GetCustomGeometryCommands(OfficeShape geometry) {
            if (geometry.Width <= 0 || geometry.Height <= 0 ||
                double.IsNaN(geometry.Width) || double.IsInfinity(geometry.Width) ||
                double.IsNaN(geometry.Height) || double.IsInfinity(geometry.Height)) {
                throw new ArgumentException(
                    "Custom geometry requires finite positive source dimensions.", nameof(geometry));
            }

            if (geometry.Kind == OfficeShapeKind.Path) {
                if (geometry.PathCommands.Count == 0) {
                    throw new ArgumentException("Custom path geometry has no commands.", nameof(geometry));
                }
                return geometry.PathCommands;
            }

            if (geometry.Kind == OfficeShapeKind.Polygon) {
                if (geometry.Points.Count < 3) {
                    throw new ArgumentException(
                        "Custom polygon geometry requires at least three points.", nameof(geometry));
                }
                var commands = new List<OfficePathCommand>(geometry.Points.Count + 2) {
                    OfficePathCommand.MoveTo(geometry.Points[0])
                };
                for (int index = 1; index < geometry.Points.Count; index++) {
                    commands.Add(OfficePathCommand.LineTo(geometry.Points[index]));
                }
                commands.Add(OfficePathCommand.Close());
                return commands;
            }

            throw new ArgumentException(
                "Custom geometry accepts shared path or polygon descriptors.", nameof(geometry));
        }

        private static A.CustomGeometry CreateCustomGeometry(OfficeShape geometry,
            IReadOnlyList<OfficePathCommand> commands) {
            var path = new A.Path {
                Width = CustomGeometryCoordinateSize,
                Height = CustomGeometryCoordinateSize
            };
            foreach (OfficePathCommand command in commands) {
                switch (command.Kind) {
                    case OfficePathCommandKind.MoveTo:
                        path.Append(new A.MoveTo(CreateCustomGeometryPoint(
                            command.Point, geometry.Width, geometry.Height)));
                        break;
                    case OfficePathCommandKind.LineTo:
                        path.Append(new A.LineTo(CreateCustomGeometryPoint(
                            command.Point, geometry.Width, geometry.Height)));
                        break;
                    case OfficePathCommandKind.QuadraticBezierTo:
                        path.Append(new A.QuadraticBezierCurveTo(
                            CreateCustomGeometryPoint(command.ControlPoint1,
                                geometry.Width, geometry.Height),
                            CreateCustomGeometryPoint(command.Point,
                                geometry.Width, geometry.Height)));
                        break;
                    case OfficePathCommandKind.CubicBezierTo:
                        path.Append(new A.CubicBezierCurveTo(
                            CreateCustomGeometryPoint(command.ControlPoint1,
                                geometry.Width, geometry.Height),
                            CreateCustomGeometryPoint(command.ControlPoint2,
                                geometry.Width, geometry.Height),
                            CreateCustomGeometryPoint(command.Point,
                                geometry.Width, geometry.Height)));
                        break;
                    case OfficePathCommandKind.Close:
                        path.Append(new A.CloseShapePath());
                        break;
                    default:
                        throw new ArgumentException(
                            "The shared path contains an unsupported command.", nameof(geometry));
                }
            }

            return new A.CustomGeometry(
                new A.AdjustValueList(),
                new A.PathList(path));
        }

        private static A.Point CreateCustomGeometryPoint(OfficePoint point,
            double sourceWidth, double sourceHeight) {
            return new A.Point {
                X = ScaleCustomGeometryCoordinate(point.X, sourceWidth),
                Y = ScaleCustomGeometryCoordinate(point.Y, sourceHeight)
            };
        }

        private static string ScaleCustomGeometryCoordinate(double value, double sourceSize) {
            double normalized = value / sourceSize
                * CustomGeometryCoordinateSize;
            double rounded = Math.Round(normalized,
                MidpointRounding.AwayFromZero);
            if (double.IsNaN(rounded) || double.IsInfinity(rounded)
                || rounded < MinimumDrawingCoordinate
                || rounded > MaximumDrawingCoordinate) {
                throw new ArgumentException(
                    "The shared path contains a coordinate that cannot be represented by DrawingML.");
            }
            long scaled = checked((long)rounded);
            return scaled.ToString(CultureInfo.InvariantCulture);
        }

        private static void ApplyCustomGeometryStyle(PowerPointAutoShape result,
            ShapeProperties properties, OfficeShape geometry) {
            properties.RemoveAllChildren<A.SolidFill>();
            properties.RemoveAllChildren<A.NoFill>();
            if (geometry.FillColor is OfficeColor fill) {
                result.FillColor = fill.ToRgbHex();
                double fillOpacity = CombineOpacity(fill.A,
                    geometry.FillOpacity);
                if (fillOpacity < 1D) {
                    result.SetFillOpacity(fillOpacity);
                }
            } else {
                A.CustomGeometry customGeometry = properties.GetFirstChild<A.CustomGeometry>()
                    ?? throw new InvalidOperationException("The added shape has no custom geometry.");
                properties.InsertAfter(new A.NoFill(), customGeometry);
            }

            result.OutlineColor = geometry.StrokeColor?.ToRgbHex();
            if (geometry.StrokeColor is OfficeColor stroke) {
                result.OutlineWidthPoints = geometry.StrokeWidth;
                result.OutlineDash = MapCustomGeometryDash(
                    geometry.StrokeDashStyle).ToOfficeEnum();
                double combinedOpacity = CombineOpacity(stroke.A,
                    geometry.StrokeOpacity);
                if (combinedOpacity < 1D) {
                    result.SetOutlineOpacity(combinedOpacity);
                }
                A.Outline outline = properties.GetFirstChild<A.Outline>()
                    ?? throw new InvalidOperationException(
                        "The added shape has no outline.");
                ApplyCustomGeometryLineCapAndJoin(outline, geometry);
                ApplyCustomGeometryLineMarkers(outline, geometry);
            }
        }

        private static A.PresetLineDashValues MapCustomGeometryDash(
            OfficeStrokeDashStyle dashStyle) => dashStyle switch {
                OfficeStrokeDashStyle.Dash => A.PresetLineDashValues.Dash,
                OfficeStrokeDashStyle.Dot => A.PresetLineDashValues.Dot,
                OfficeStrokeDashStyle.DashDot => A.PresetLineDashValues.DashDot,
                OfficeStrokeDashStyle.DashDotDot =>
                    A.PresetLineDashValues.LargeDashDotDot,
                _ => A.PresetLineDashValues.Solid
            };

        private static void ApplyCustomGeometryLineCapAndJoin(
            A.Outline outline, OfficeShape geometry) {
            outline.CapType = geometry.StrokeLineCap switch {
                OfficeStrokeLineCap.Round => A.LineCapValues.Round,
                OfficeStrokeLineCap.Square => A.LineCapValues.Square,
                OfficeStrokeLineCap.Butt => A.LineCapValues.Flat,
                _ => (A.LineCapValues?)null
            };
            outline.RemoveAllChildren<A.Round>();
            outline.RemoveAllChildren<A.LineJoinBevel>();
            outline.RemoveAllChildren<A.Miter>();
            if (geometry.StrokeLineJoin == OfficeStrokeLineJoin.Round) {
                outline.Append(new A.Round());
            } else if (geometry.StrokeLineJoin == OfficeStrokeLineJoin.Bevel) {
                outline.Append(new A.LineJoinBevel());
            } else if (geometry.StrokeLineJoin == OfficeStrokeLineJoin.Miter) {
                outline.Append(new A.Miter());
            }
        }

        private static void ApplyCustomGeometryLineMarkers(A.Outline outline,
            OfficeShape geometry) {
            outline.RemoveAllChildren<A.HeadEnd>();
            outline.RemoveAllChildren<A.TailEnd>();
            if (geometry.StrokeStartMarker != null) {
                outline.Append(CreateCustomGeometryLineEnd<A.HeadEnd>(
                    geometry.StrokeStartMarker, geometry.StrokeWidth));
            }
            if (geometry.StrokeEndMarker != null) {
                outline.Append(CreateCustomGeometryLineEnd<A.TailEnd>(
                    geometry.StrokeEndMarker, geometry.StrokeWidth));
            }
        }

        private static T CreateCustomGeometryLineEnd<T>(OfficeLineMarker marker,
            double strokeWidth) where T : A.LineEndPropertiesType, new() {
            double width = marker.Width / Math.Max(strokeWidth, 0.01D);
            double length = marker.Length / Math.Max(strokeWidth, 0.01D);
            return new T {
                Type = marker.Kind switch {
                    OfficeLineMarkerKind.Triangle => A.LineEndValues.Triangle,
                    OfficeLineMarkerKind.Stealth => A.LineEndValues.Stealth,
                    OfficeLineMarkerKind.Diamond => A.LineEndValues.Diamond,
                    OfficeLineMarkerKind.Oval => A.LineEndValues.Oval,
                    OfficeLineMarkerKind.Arrow => A.LineEndValues.Arrow,
                    _ => A.LineEndValues.None
                },
                Width = width < 3.75D ? A.LineEndWidthValues.Small
                    : width > 5.25D ? A.LineEndWidthValues.Large
                    : A.LineEndWidthValues.Medium,
                Length = length < 5D ? A.LineEndLengthValues.Small
                    : length > 7D ? A.LineEndLengthValues.Large
                    : A.LineEndLengthValues.Medium
            };
        }

        private static double CombineOpacity(byte colorAlpha,
            double? declaredOpacity) {
            double opacity = declaredOpacity ?? 1D;
            double clampedOpacity = opacity < 0D ? 0D
                : opacity > 1D ? 1D
                : opacity;
            return colorAlpha / (double)byte.MaxValue * clampedOpacity;
        }

        private static void ValidateCustomGeometryOpacity(double? opacity,
            string propertyName) {
            if (opacity.HasValue && (double.IsNaN(opacity.Value)
                    || double.IsInfinity(opacity.Value)
                    || opacity.Value < 0D || opacity.Value > 1D)) {
                throw new ArgumentException(
                    $"Custom geometry {propertyName} must be finite and between 0 and 1.",
                    "geometry");
            }
        }

        private static void ValidateCustomGeometryTransformCoordinate(long value,
            string parameterName) {
            PowerPointDrawingValueValidator.ValidateCoordinate(value,
                parameterName, "Custom geometry coordinates");
        }

        private static void ValidateCustomGeometryExtent(long value,
            string parameterName) {
            if (value > MaximumDrawingCoordinate) {
                throw new ArgumentOutOfRangeException(parameterName,
                    $"Custom geometry extents must not exceed {MaximumDrawingCoordinate} EMUs.");
            }
        }
    }
}
