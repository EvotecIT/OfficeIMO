using System;
using System.Collections.Generic;
using System.Globalization;
using DocumentFormat.OpenXml.Presentation;
using OfficeIMO.Drawing;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.PowerPoint {
    public partial class PowerPointSlide {
        private const long CustomGeometryCoordinateSize = 100000L;
        private const int MaximumCustomGeometryCommands = 20000;

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

            IReadOnlyList<OfficePathCommand> commands = GetCustomGeometryCommands(geometry);
            if (commands.Count > MaximumCustomGeometryCommands) {
                throw new ArgumentException(
                    $"Custom geometry supports at most {MaximumCustomGeometryCommands} path commands.",
                    nameof(geometry));
            }

            PowerPointAutoShape result = AddShape(
                A.ShapeTypeValues.Rectangle, left, top, width, height, name);
            Shape shape = (Shape)result.Element;
            ShapeProperties properties = shape.ShapeProperties
                ?? throw new InvalidOperationException("The added shape has no shape properties.");
            A.Transform2D transform = properties.GetFirstChild<A.Transform2D>()
                ?? throw new InvalidOperationException("The added shape has no transform.");
            properties.RemoveAllChildren<A.PresetGeometry>();
            properties.InsertAfter(CreateCustomGeometry(geometry, commands), transform);
            ApplyCustomGeometryStyle(result, properties, geometry);
            return result;
        }

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
            long scaled = (long)Math.Round(
                value / sourceSize * CustomGeometryCoordinateSize,
                MidpointRounding.AwayFromZero);
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
                double combinedOpacity = CombineOpacity(stroke.A,
                    geometry.StrokeOpacity);
                if (combinedOpacity < 1D) {
                    result.SetOutlineOpacity(combinedOpacity);
                }
            }
        }

        private static double CombineOpacity(byte colorAlpha,
            double? declaredOpacity) {
            double opacity = declaredOpacity ?? 1D;
            double clampedOpacity = opacity < 0D ? 0D
                : opacity > 1D ? 1D
                : opacity;
            return colorAlpha / (double)byte.MaxValue * clampedOpacity;
        }
    }
}
