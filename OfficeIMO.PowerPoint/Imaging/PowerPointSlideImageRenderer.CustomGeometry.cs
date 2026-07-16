using System;
using System.Collections.Generic;
using System.Globalization;
using OfficeIMO.Drawing;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.PowerPoint {
    internal static partial class PowerPointSlideImageRenderer {
        private const double DrawingMlAngleUnitsPerDegree = 60000D;

        private static bool TryAddCustomGeometryShape(
            OfficeDrawing drawing,
            PowerPointShape shape,
            double left,
            double top,
            double width,
            double height,
            List<OfficeImageExportDiagnostic> diagnostics,
            PowerPointShapeBoundsMapping mapping,
            A.ColorScheme? colorScheme) {
            A.CustomGeometry? customGeometry = GetOpenXmlShapeProperties(shape)?.GetFirstChild<A.CustomGeometry>();
            if (customGeometry == null) {
                return false;
            }

            if (!TryCreateCustomGeometryShape(customGeometry, width, height, out OfficeShape? drawingShape) || drawingShape == null) {
                AddUnsupportedShapeDiagnostic(diagnostics, shape, "Skipped a PowerPoint custom geometry shape because its paths use commands, guides, or formulas that are not yet projected through OfficeIMO.Drawing.");
                return true;
            }

            ApplyShapeStyle(drawingShape, shape, colorScheme, mapping, diagnostics);
            ApplyShapeTransform(drawingShape, shape, width, height);
            drawing.AddShape(drawingShape, left, top);
            return true;
        }

        private static bool TryCreateCustomGeometryShape(A.CustomGeometry customGeometry, double width, double height, out OfficeShape? drawingShape) {
            drawingShape = null;
            A.PathList? pathList = customGeometry.PathList;
            if (pathList == null) {
                return false;
            }

            var commands = new List<OfficePathCommand>();
            foreach (A.Path path in pathList.Elements<A.Path>()) {
                if (!TryAppendCustomGeometryPath(customGeometry, path, width, height, commands)) {
                    return false;
                }
            }

            if (commands.Count == 0) {
                return false;
            }

            try {
                drawingShape = OfficeShape.Path(commands);
                return true;
            } catch (ArgumentException) {
                drawingShape = null;
                return false;
            }
        }

        private static bool TryAppendCustomGeometryPath(A.CustomGeometry customGeometry, A.Path path, double width, double height, List<OfficePathCommand> commands) {
            long coordinateWidth = path.Width?.Value ?? 0L;
            long coordinateHeight = path.Height?.Value ?? 0L;
            if (coordinateWidth <= 0L || coordinateHeight <= 0L) {
                return false;
            }

            if (!TryCreateCustomGeometryGuideContext(customGeometry, coordinateWidth, coordinateHeight, out Dictionary<string, double>? guides) || guides == null) {
                return false;
            }

            double scaleX = width / coordinateWidth;
            double scaleY = height / coordinateHeight;
            bool hasMove = false;
            bool hasDraw = false;
            OfficePoint currentPoint = default;

            foreach (DocumentFormat.OpenXml.OpenXmlElement child in path.ChildElements) {
                if (child is A.MoveTo moveTo) {
                    if (!TryGetCustomGeometryPoint(moveTo.Point, scaleX, scaleY, guides, out OfficePoint point)) {
                        return false;
                    }

                    commands.Add(OfficePathCommand.MoveTo(point));
                    currentPoint = point;
                    hasMove = true;
                } else if (child is A.LineTo lineTo) {
                    if (!hasMove || !TryGetCustomGeometryPoint(lineTo.Point, scaleX, scaleY, guides, out OfficePoint point)) {
                        return false;
                    }

                    commands.Add(OfficePathCommand.LineTo(point));
                    currentPoint = point;
                    hasDraw = true;
                } else if (child is A.QuadraticBezierCurveTo quadraticBezier) {
                    if (!hasMove || !TryGetCustomGeometryPoints(quadraticBezier, 2, scaleX, scaleY, guides, out List<OfficePoint>? points)) {
                        return false;
                    }

                    commands.Add(OfficePathCommand.QuadraticBezierTo(points![0], points[1]));
                    currentPoint = points[1];
                    hasDraw = true;
                } else if (child is A.CubicBezierCurveTo cubicBezier) {
                    if (!hasMove || !TryGetCustomGeometryPoints(cubicBezier, 3, scaleX, scaleY, guides, out List<OfficePoint>? points)) {
                        return false;
                    }

                    commands.Add(OfficePathCommand.CubicBezierTo(points![0], points[1], points[2]));
                    currentPoint = points[2];
                    hasDraw = true;
                } else if (child is A.ArcTo arcTo) {
                    if (!hasMove || !TryCreateCustomGeometryArcCommands(arcTo, currentPoint, scaleX, scaleY, guides, out List<OfficePathCommand>? arcCommands)) {
                        return false;
                    }

                    List<OfficePathCommand> resolvedArcCommands = arcCommands!;
                    commands.AddRange(resolvedArcCommands);
                    currentPoint = resolvedArcCommands[resolvedArcCommands.Count - 1].Point;
                    hasDraw = true;
                } else if (child is A.CloseShapePath) {
                    if (!hasMove || !hasDraw) {
                        return false;
                    }

                    commands.Add(OfficePathCommand.Close());
                } else {
                    return false;
                }
            }

            return hasMove && hasDraw;
        }

        private static bool TryCreateCustomGeometryArcCommands(A.ArcTo arcTo, OfficePoint currentPoint, double scaleX, double scaleY, Dictionary<string, double> guides, out List<OfficePathCommand>? commands) {
            commands = null;
            if (!TryResolveCustomGeometryCoordinate(arcTo.WidthRadius?.Value, guides, out double widthRadius) ||
                !TryResolveCustomGeometryCoordinate(arcTo.HeightRadius?.Value, guides, out double heightRadius) ||
                !TryResolveCustomGeometryCoordinate(arcTo.StartAngle?.Value, guides, out double startAngle) ||
                !TryResolveCustomGeometryCoordinate(arcTo.SwingAngle?.Value, guides, out double swingAngle)) {
                return false;
            }

            double scaledRadiusX = widthRadius * scaleX;
            double scaledRadiusY = heightRadius * scaleY;
            if (scaledRadiusX <= 0D || scaledRadiusY <= 0D) {
                return false;
            }

            double startRadians = OfficeGeometry.DegreesToRadians(startAngle / DrawingMlAngleUnitsPerDegree);
            double sweepRadians = OfficeGeometry.DegreesToRadians(swingAngle / DrawingMlAngleUnitsPerDegree);
            try {
                commands = OfficeGeometry.CreateEllipticalArcCubicBezierCommands(currentPoint, scaledRadiusX, scaledRadiusY, startRadians, sweepRadians);
            } catch (ArgumentOutOfRangeException) {
                commands = null;
                return false;
            }

            return commands.Count > 0;
        }

        private static bool TryGetCustomGeometryPoints(DocumentFormat.OpenXml.OpenXmlElement element, int expectedCount, double scaleX, double scaleY, Dictionary<string, double> guides, out List<OfficePoint>? points) {
            points = new List<OfficePoint>(expectedCount);
            foreach (A.Point point in element.Elements<A.Point>()) {
                if (!TryGetCustomGeometryPoint(point, scaleX, scaleY, guides, out OfficePoint officePoint)) {
                    points = null;
                    return false;
                }

                points.Add(officePoint);
            }

            if (points.Count != expectedCount) {
                points = null;
                return false;
            }

            return true;
        }

        private static bool TryGetCustomGeometryPoint(A.Point? point, double scaleX, double scaleY, Dictionary<string, double> guides, out OfficePoint officePoint) {
            officePoint = default;
            if (point == null ||
                !TryResolveCustomGeometryCoordinate(point.X?.Value, guides, out double x) ||
                !TryResolveCustomGeometryCoordinate(point.Y?.Value, guides, out double y)) {
                return false;
            }

            officePoint = new OfficePoint(x * scaleX, y * scaleY);
            return true;
        }

        private static bool TryCreateCustomGeometryGuideContext(A.CustomGeometry customGeometry, long coordinateWidth, long coordinateHeight, out Dictionary<string, double>? guides) {
            guides = new Dictionary<string, double>(StringComparer.Ordinal) {
                ["l"] = 0D,
                ["t"] = 0D,
                ["r"] = coordinateWidth,
                ["b"] = coordinateHeight,
                ["w"] = coordinateWidth,
                ["h"] = coordinateHeight,
                ["hc"] = coordinateWidth / 2D,
                ["vc"] = coordinateHeight / 2D,
                ["wd2"] = coordinateWidth / 2D,
                ["wd3"] = coordinateWidth / 3D,
                ["wd4"] = coordinateWidth / 4D,
                ["wd5"] = coordinateWidth / 5D,
                ["wd6"] = coordinateWidth / 6D,
                ["wd8"] = coordinateWidth / 8D,
                ["wd10"] = coordinateWidth / 10D,
                ["hd2"] = coordinateHeight / 2D,
                ["hd3"] = coordinateHeight / 3D,
                ["hd4"] = coordinateHeight / 4D,
                ["hd5"] = coordinateHeight / 5D,
                ["hd6"] = coordinateHeight / 6D,
                ["hd8"] = coordinateHeight / 8D,
                ["hd10"] = coordinateHeight / 10D
            };
            double shortSide = Math.Min(coordinateWidth, coordinateHeight);
            guides["ss"] = shortSide;
            guides["ls"] = Math.Max(coordinateWidth, coordinateHeight);
            guides["ssd2"] = shortSide / 2D;
            guides["ssd4"] = shortSide / 4D;
            guides["ssd6"] = shortSide / 6D;
            guides["ssd8"] = shortSide / 8D;
            guides["ssd16"] = shortSide / 16D;
            guides["ssd32"] = shortSide / 32D;

            if (!TryEvaluateCustomGeometryGuideList(customGeometry.AdjustValueList, guides) ||
                !TryEvaluateCustomGeometryGuideList(customGeometry.ShapeGuideList, guides)) {
                guides = null;
                return false;
            }

            return true;
        }

        private static bool TryEvaluateCustomGeometryGuideList(DocumentFormat.OpenXml.OpenXmlElement? guideList, Dictionary<string, double> guides) {
            if (guideList == null) {
                return true;
            }

            foreach (A.ShapeGuide guide in guideList.Elements<A.ShapeGuide>()) {
                string? name = guide.Name?.Value;
                string? formula = guide.Formula?.Value;
                if (string.IsNullOrWhiteSpace(name) ||
                    string.IsNullOrWhiteSpace(formula) ||
                    !TryEvaluateCustomGeometryGuideFormula(formula!, guides, out double value)) {
                    return false;
                }

                guides[name!] = value;
            }

            return true;
        }

        private static bool TryEvaluateCustomGeometryGuideFormula(string formula, Dictionary<string, double> guides, out double value) {
            value = 0D;
            string[] parts = formula.Split(new[] { ' ', '\t', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
            if (parts.Length == 0) {
                return false;
            }

            switch (parts[0]) {
                case "val":
                    return parts.Length == 2 && TryResolveCustomGeometryCoordinate(parts[1], guides, out value);
                case "+-":
                    if (parts.Length == 4 &&
                        TryResolveCustomGeometryCoordinate(parts[1], guides, out double addend1) &&
                        TryResolveCustomGeometryCoordinate(parts[2], guides, out double addend2) &&
                        TryResolveCustomGeometryCoordinate(parts[3], guides, out double subtractend)) {
                        value = addend1 + addend2 - subtractend;
                        return IsFiniteCustomGeometryValue(value);
                    }

                    return false;
                case "+/":
                    if (parts.Length == 4 &&
                        TryResolveCustomGeometryCoordinate(parts[1], guides, out double dividend1) &&
                        TryResolveCustomGeometryCoordinate(parts[2], guides, out double dividend2) &&
                        TryResolveCustomGeometryCoordinate(parts[3], guides, out double divisor) &&
                        Math.Abs(divisor) > double.Epsilon) {
                        value = (dividend1 + dividend2) / divisor;
                        return IsFiniteCustomGeometryValue(value);
                    }

                    return false;
                case "*/":
                    if (parts.Length == 4 &&
                        TryResolveCustomGeometryCoordinate(parts[1], guides, out double multiplicand) &&
                        TryResolveCustomGeometryCoordinate(parts[2], guides, out double multiplier) &&
                        TryResolveCustomGeometryCoordinate(parts[3], guides, out double denominator) &&
                        Math.Abs(denominator) > double.Epsilon) {
                        value = multiplicand * multiplier / denominator;
                        return IsFiniteCustomGeometryValue(value);
                    }

                    return false;
                case "min":
                    if (parts.Length == 3 &&
                        TryResolveCustomGeometryCoordinate(parts[1], guides, out double min1) &&
                        TryResolveCustomGeometryCoordinate(parts[2], guides, out double min2)) {
                        value = Math.Min(min1, min2);
                        return true;
                    }

                    return false;
                case "max":
                    if (parts.Length == 3 &&
                        TryResolveCustomGeometryCoordinate(parts[1], guides, out double max1) &&
                        TryResolveCustomGeometryCoordinate(parts[2], guides, out double max2)) {
                        value = Math.Max(max1, max2);
                        return true;
                    }

                    return false;
                case "abs":
                    if (parts.Length == 2 && TryResolveCustomGeometryCoordinate(parts[1], guides, out double absolute)) {
                        value = Math.Abs(absolute);
                        return true;
                    }

                    return false;
                case "sqrt":
                    if (parts.Length == 2 &&
                        TryResolveCustomGeometryCoordinate(parts[1], guides, out double squareRoot) &&
                        squareRoot >= 0D) {
                        value = Math.Sqrt(squareRoot);
                        return IsFiniteCustomGeometryValue(value);
                    }

                    return false;
                case "sin":
                    if (parts.Length == 3 &&
                        TryResolveCustomGeometryCoordinate(parts[1], guides, out double sineMagnitude) &&
                        TryResolveCustomGeometryCoordinate(parts[2], guides, out double sineAngle)) {
                        value = sineMagnitude * Math.Sin(ConvertCustomGeometryAngleToRadians(sineAngle));
                        return IsFiniteCustomGeometryValue(value);
                    }

                    return false;
                case "cos":
                    if (parts.Length == 3 &&
                        TryResolveCustomGeometryCoordinate(parts[1], guides, out double cosineMagnitude) &&
                        TryResolveCustomGeometryCoordinate(parts[2], guides, out double cosineAngle)) {
                        value = cosineMagnitude * Math.Cos(ConvertCustomGeometryAngleToRadians(cosineAngle));
                        return IsFiniteCustomGeometryValue(value);
                    }

                    return false;
                case "tan":
                    if (parts.Length == 3 &&
                        TryResolveCustomGeometryCoordinate(parts[1], guides, out double tangentMagnitude) &&
                        TryResolveCustomGeometryCoordinate(parts[2], guides, out double tangentAngle)) {
                        value = tangentMagnitude * Math.Tan(ConvertCustomGeometryAngleToRadians(tangentAngle));
                        return IsFiniteCustomGeometryValue(value);
                    }

                    return false;
                case "mod":
                    if (parts.Length == 4 &&
                        TryResolveCustomGeometryCoordinate(parts[1], guides, out double modulusX) &&
                        TryResolveCustomGeometryCoordinate(parts[2], guides, out double modulusY) &&
                        TryResolveCustomGeometryCoordinate(parts[3], guides, out double modulusZ)) {
                        value = Math.Sqrt((modulusX * modulusX) + (modulusY * modulusY) + (modulusZ * modulusZ));
                        return IsFiniteCustomGeometryValue(value);
                    }

                    return false;
                case "at2":
                    if (parts.Length == 3 &&
                        TryResolveCustomGeometryCoordinate(parts[1], guides, out double arcTangentY) &&
                        TryResolveCustomGeometryCoordinate(parts[2], guides, out double arcTangentX)) {
                        value = ConvertRadiansToCustomGeometryAngle(Math.Atan2(arcTangentY, arcTangentX));
                        return IsFiniteCustomGeometryValue(value);
                    }

                    return false;
                case "cat2":
                    if (parts.Length == 4 &&
                        TryResolveCustomGeometryCoordinate(parts[1], guides, out double cosineAtanMagnitude) &&
                        TryResolveCustomGeometryCoordinate(parts[2], guides, out double cosineAtanX) &&
                        TryResolveCustomGeometryCoordinate(parts[3], guides, out double cosineAtanY)) {
                        value = cosineAtanMagnitude * Math.Cos(Math.Atan2(cosineAtanY, cosineAtanX));
                        return IsFiniteCustomGeometryValue(value);
                    }

                    return false;
                case "sat2":
                    if (parts.Length == 4 &&
                        TryResolveCustomGeometryCoordinate(parts[1], guides, out double sineAtanMagnitude) &&
                        TryResolveCustomGeometryCoordinate(parts[2], guides, out double sineAtanX) &&
                        TryResolveCustomGeometryCoordinate(parts[3], guides, out double sineAtanY)) {
                        value = sineAtanMagnitude * Math.Sin(Math.Atan2(sineAtanY, sineAtanX));
                        return IsFiniteCustomGeometryValue(value);
                    }

                    return false;
                case "pin":
                    if (parts.Length == 4 &&
                        TryResolveCustomGeometryCoordinate(parts[1], guides, out double low) &&
                        TryResolveCustomGeometryCoordinate(parts[2], guides, out double current) &&
                        TryResolveCustomGeometryCoordinate(parts[3], guides, out double high)) {
                        value = Math.Max(low, Math.Min(current, high));
                        return IsFiniteCustomGeometryValue(value);
                    }

                    return false;
                case "?:":
                    if (parts.Length == 4 &&
                        TryResolveCustomGeometryCoordinate(parts[1], guides, out double condition) &&
                        TryResolveCustomGeometryCoordinate(parts[2], guides, out double trueValue) &&
                        TryResolveCustomGeometryCoordinate(parts[3], guides, out double falseValue)) {
                        value = condition > 0D ? trueValue : falseValue;
                        return IsFiniteCustomGeometryValue(value);
                    }

                    return false;
                default:
                    return false;
            }
        }

        private static bool TryResolveCustomGeometryCoordinate(string? value, Dictionary<string, double> guides, out double coordinate) {
            if (TryParseCustomGeometryCoordinate(value, out coordinate)) {
                return true;
            }

            return !string.IsNullOrWhiteSpace(value) && guides.TryGetValue(value!, out coordinate);
        }

        private static bool TryParseCustomGeometryCoordinate(string? value, out double coordinate) {
            coordinate = 0D;
            return !string.IsNullOrWhiteSpace(value) &&
                double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out coordinate) &&
                !double.IsNaN(coordinate) &&
                !double.IsInfinity(coordinate);
        }

        private static double ConvertCustomGeometryAngleToRadians(double angle) =>
            OfficeGeometry.DegreesToRadians(angle / DrawingMlAngleUnitsPerDegree);

        private static double ConvertRadiansToCustomGeometryAngle(double radians) =>
            radians * 180D / Math.PI * DrawingMlAngleUnitsPerDegree;

        private static bool IsFiniteCustomGeometryValue(double value) =>
            !double.IsNaN(value) && !double.IsInfinity(value);
    }
}
