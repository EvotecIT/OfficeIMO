using System.Globalization;
using OfficeIMO.Drawing;
using OfficeIMO.OpenXml.Internal;

namespace OfficeIMO.PowerPoint {
    internal static partial class PowerPointSlideImageRenderer {
        private const double DrawingMlAngleUnitsPerDegree = 60000D;

        internal static bool TryCreateCustomGeometryDrawing(
            PowerPointShape shape, double width, double height,
            out OfficeDrawing? drawing, out string? reason) {
            ArgumentNullException.ThrowIfNull(shape);
            drawing = null;
            reason = null;
            if (width <= 0D || height <= 0D) {
                reason = "The custom-geometry frame has no renderable area.";
                return false;
            }
            var candidate = new OfficeDrawing(width, height);
            var diagnostics = new List<OfficeImageExportDiagnostic>();
            OfficeXmlNode colorScheme = shape.OwnerSlide == null
                ? default
                : GetSlideColorScheme(shape.OwnerSlide);
            if (!TryAddCustomGeometryShape(candidate, shape,
                    left: 0D, top: 0D, width, height, diagnostics,
                    PowerPointShapeBoundsMapping.Identity, colorScheme)) {
                reason = "The PowerPoint shape has no custom geometry.";
                return false;
            }
            OfficeImageExportDiagnostic? failure = diagnostics.FirstOrDefault(
                diagnostic => diagnostic.Severity
                        != OfficeImageExportDiagnosticSeverity.Info
                    || diagnostic.LossKind != OfficeConversionLossKind.None);
            if (failure != null || candidate.Elements.Count == 0) {
                reason = failure?.Message
                    ?? "The custom geometry produced no visible Drawing paths.";
                return false;
            }
            drawing = candidate;
            return true;
        }

        private static bool TryAddCustomGeometryShape(
            OfficeDrawing drawing,
            PowerPointShape shape,
            double left,
            double top,
            double width,
            double height,
            List<OfficeImageExportDiagnostic> diagnostics,
            PowerPointShapeBoundsMapping mapping,
            OfficeXmlNode colorScheme) {
            OfficeXmlNode customGeometry = GetShapePropertiesElement(shape).Child(D, "custGeom");
            if (customGeometry.IsNone) {
                return false;
            }

            if (!TryCreateCustomGeometryShapes(customGeometry, width, height,
                    out List<CustomGeometryPathProjection>? projections)
                    || projections == null) {
                AddUnsupportedShapeDiagnostic(diagnostics, shape, "Skipped a PowerPoint custom geometry shape because its paths use commands, guides, formulas, or per-path styling that are not yet projected faithfully through OfficeIMO.Drawing.");
                return true;
            }

            int visiblePathCount = projections.Count(projection =>
                projection.HasFill || projection.HasStroke);
            if (visiblePathCount > 1
                && HasProjectableShapeEffects(shape, colorScheme, mapping)) {
                AddUnsupportedShapeDiagnostic(diagnostics, shape,
                    "Skipped a multi-path PowerPoint custom geometry shape because its shape-level shadow or glow cannot be composited faithfully across separate Drawing paths.");
                return true;
            }

            foreach (CustomGeometryPathProjection projection in projections) {
                OfficeShape drawingShape = projection.Shape;
                ApplyShapeStyle(drawingShape, shape, colorScheme, mapping,
                    diagnostics);
                if (!projection.HasFill) {
                    drawingShape.FillColor = null;
                    drawingShape.FillGradient = null;
                    drawingShape.FillRadialGradient = null;
                }
                if (!projection.HasStroke) {
                    drawingShape.StrokeColor = null;
                    drawingShape.StrokeGradient = null;
                    drawingShape.StrokeRadialGradient = null;
                    drawingShape.StrokeWidth = 0D;
                    drawingShape.StrokeStartMarker = null;
                    drawingShape.StrokeEndMarker = null;
                }
                if (!projection.HasFill && !projection.HasStroke) {
                    drawingShape.Glow = null;
                    drawingShape.Shadow = null;
                }
                ApplyShapeTransformForExport(drawingShape, shape, width, height);
                drawing.AddShape(drawingShape, left, top);
            }
            return true;
        }

        private static bool HasProjectableShapeEffects(PowerPointShape shape,
            OfficeXmlNode colorScheme,
            PowerPointShapeBoundsMapping mapping) {
            OfficeXmlNode effects = GetShapePropertiesElement(shape).Child(D, "effectLst");
            if (effects.IsNone) return false;

            OfficeXmlNode glow = effects.Child(D, "glow");
            if (!glow.IsNone && TryCreateGlow(glow, colorScheme, mapping,
                    out _)) {
                return true;
            }

            OfficeXmlNode shadow = effects.Child(D, "outerShdw");
            return !shadow.IsNone && TryCreateShadow(shadow, colorScheme,
                mapping, out _);
        }

        private static bool TryCreateCustomGeometryShapes(
            OfficeXmlNode customGeometry, double width, double height,
            out List<CustomGeometryPathProjection>? projections) {
            projections = null;
            OfficeXmlNode pathList = customGeometry.Child(D, "pathLst");
            if (pathList.IsNone) {
                return false;
            }

            var resolved = new List<CustomGeometryPathProjection>();
            foreach (OfficeXmlNode path in pathList.Children(D, "path")) {
                // DrawingML ST_PathFillMode tokens: only no fill and normal fill project.
                string? fillMode = path.Attribute("fill");
                if (fillMode != null
                    && fillMode != "none"
                    && fillMode != "norm") {
                    return false;
                }

                var commands = new List<OfficePathCommand>();
                if (!TryAppendCustomGeometryPath(customGeometry, path, width, height, commands)) {
                    return false;
                }

                try {
                    OfficeShape projectedShape = OfficeShape.Path(
                        width, height, commands);
                    projectedShape.FillRule = OfficeFillRule.NonZero;
                    resolved.Add(new CustomGeometryPathProjection(
                        projectedShape,
                        fillMode != "none",
                        path.BooleanAttribute("stroke") != false));
                } catch (ArgumentException) {
                    return false;
                }
            }

            if (resolved.Count == 0) return false;
            projections = resolved;
            return true;
        }

        private sealed class CustomGeometryPathProjection {
            internal CustomGeometryPathProjection(OfficeShape shape,
                bool hasFill, bool hasStroke) {
                Shape = shape;
                HasFill = hasFill;
                HasStroke = hasStroke;
            }

            internal OfficeShape Shape { get; }
            internal bool HasFill { get; }
            internal bool HasStroke { get; }
        }

        private static bool TryAppendCustomGeometryPath(OfficeXmlNode customGeometry, OfficeXmlNode path, double width, double height, List<OfficePathCommand> commands) {
            long coordinateWidth = path.Int64Attribute("w") ?? 0L;
            long coordinateHeight = path.Int64Attribute("h") ?? 0L;
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
            OfficePoint currentSubpathStart = default;

            foreach (OfficeXmlNode child in path.Children()) {
                string command = child.NamespaceId == D ? child.LocalName : string.Empty;
                if (command == "moveTo") {
                    if (!TryGetCustomGeometryPoint(child.Child(D, "pt"), scaleX, scaleY, guides, out OfficePoint point)) {
                        return false;
                    }

                    commands.Add(OfficePathCommand.MoveTo(point));
                    currentPoint = point;
                    currentSubpathStart = point;
                    hasMove = true;
                } else if (command == "lnTo") {
                    if (!hasMove || !TryGetCustomGeometryPoint(child.Child(D, "pt"), scaleX, scaleY, guides, out OfficePoint point)) {
                        return false;
                    }

                    commands.Add(OfficePathCommand.LineTo(point));
                    currentPoint = point;
                    hasDraw = true;
                } else if (command == "quadBezTo") {
                    if (!hasMove || !TryGetCustomGeometryPoints(child, 2, scaleX, scaleY, guides, out List<OfficePoint>? points)) {
                        return false;
                    }

                    commands.Add(OfficePathCommand.QuadraticBezierTo(points![0], points[1]));
                    currentPoint = points[1];
                    hasDraw = true;
                } else if (command == "cubicBezTo") {
                    if (!hasMove || !TryGetCustomGeometryPoints(child, 3, scaleX, scaleY, guides, out List<OfficePoint>? points)) {
                        return false;
                    }

                    commands.Add(OfficePathCommand.CubicBezierTo(points![0], points[1], points[2]));
                    currentPoint = points[2];
                    hasDraw = true;
                } else if (command == "arcTo") {
                    if (!hasMove || !TryCreateCustomGeometryArcCommands(child, currentPoint, scaleX, scaleY, guides, out List<OfficePathCommand>? arcCommands)) {
                        return false;
                    }

                    List<OfficePathCommand> resolvedArcCommands = arcCommands!;
                    commands.AddRange(resolvedArcCommands);
                    currentPoint = resolvedArcCommands[resolvedArcCommands.Count - 1].Point;
                    hasDraw = true;
                } else if (command == "close") {
                    if (!hasMove || !hasDraw) {
                        return false;
                    }

                    commands.Add(OfficePathCommand.Close());
                    currentPoint = currentSubpathStart;
                } else {
                    return false;
                }
            }

            return hasMove && hasDraw;
        }

        private static bool TryCreateCustomGeometryArcCommands(OfficeXmlNode arcTo, OfficePoint currentPoint, double scaleX, double scaleY, Dictionary<string, double> guides, out List<OfficePathCommand>? commands) {
            commands = null;
            if (!TryResolveCustomGeometryCoordinate(arcTo.Attribute("wR"), guides, out double widthRadius) ||
                !TryResolveCustomGeometryCoordinate(arcTo.Attribute("hR"), guides, out double heightRadius) ||
                !TryResolveCustomGeometryCoordinate(arcTo.Attribute("stAng"), guides, out double startAngle) ||
                !TryResolveCustomGeometryCoordinate(arcTo.Attribute("swAng"), guides, out double swingAngle)) {
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

        private static bool TryGetCustomGeometryPoints(OfficeXmlNode element, int expectedCount, double scaleX, double scaleY, Dictionary<string, double> guides, out List<OfficePoint>? points) {
            points = new List<OfficePoint>(expectedCount);
            foreach (OfficeXmlNode point in element.Children(D, "pt")) {
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

        private static bool TryGetCustomGeometryPoint(OfficeXmlNode point, double scaleX, double scaleY, Dictionary<string, double> guides, out OfficePoint officePoint) {
            officePoint = default;
            if (point.IsNone ||
                !TryResolveCustomGeometryCoordinate(point.Attribute("x"), guides, out double x) ||
                !TryResolveCustomGeometryCoordinate(point.Attribute("y"), guides, out double y)) {
                return false;
            }

            officePoint = new OfficePoint(x * scaleX, y * scaleY);
            return true;
        }

        private static bool TryCreateCustomGeometryGuideContext(OfficeXmlNode customGeometry, long coordinateWidth, long coordinateHeight, out Dictionary<string, double>? guides) {
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

            if (!TryEvaluateCustomGeometryGuideList(customGeometry.Child(D, "avLst"), guides) ||
                !TryEvaluateCustomGeometryGuideList(customGeometry.Child(D, "gdLst"), guides)) {
                guides = null;
                return false;
            }

            return true;
        }

        private static bool TryEvaluateCustomGeometryGuideList(OfficeXmlNode guideList, Dictionary<string, double> guides) {
            if (guideList.IsNone) {
                return true;
            }

            foreach (OfficeXmlNode guide in guideList.Children(D, "gd")) {
                string? name = guide.Attribute("name");
                string? formula = guide.Attribute("fmla");
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
                        TryResolveCustomGeometryCoordinate(parts[1], guides, out double arcTangentX) &&
                        TryResolveCustomGeometryCoordinate(parts[2], guides, out double arcTangentY)) {
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
                OfficeIMO.Core.Internal.OfficeRoundTripText.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out coordinate) &&
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
