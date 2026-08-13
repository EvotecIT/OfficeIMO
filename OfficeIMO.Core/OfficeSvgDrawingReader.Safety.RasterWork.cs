using System;
using System.Collections.Generic;
using System.Xml.Linq;

namespace OfficeIMO.Drawing;

public static partial class OfficeSvgDrawingReader {
    private sealed class SvgRasterWorkBudget {
        private const double MaximumViewportRepaints = 256D;
        private readonly double _viewLeft;
        private readonly double _viewTop;
        private readonly double _viewRight;
        private readonly double _viewBottom;
        private readonly double _pixelScaleX;
        private readonly double _pixelScaleY;
        private readonly double _viewportPixels;
        private double _remainingWork;

        internal SvgRasterWorkBudget(
            double maximumViewportPixels,
            double viewX,
            double viewY,
            double viewWidth,
            double viewHeight,
            double viewportWidth,
            double viewportHeight) {
            _viewLeft = viewX;
            _viewTop = viewY;
            _viewRight = viewX + viewWidth;
            _viewBottom = viewY + viewHeight;
            _pixelScaleX = viewportWidth / viewWidth;
            _pixelScaleY = viewportHeight / viewHeight;
            _viewportPixels = Math.Min(maximumViewportPixels, viewportWidth * viewportHeight);
            _remainingWork = _viewportPixels * MaximumViewportRepaints;
        }

        internal bool TryChargeRenderedElement(XElement element, OfficeTransform transform) {
            string name = element.Name.LocalName.ToLowerInvariant();
            if (name is "svg" or "g" or "a" or "switch" or "symbol" or "pattern" or "mask" or
                "clippath" or "filter" or "marker" or "style" or "use" or "defs" or "title" or
                "desc" or "metadata" or "lineargradient" or "radialgradient" or "stop") return true;

            if (!TryGetPaintBounds(element, name, out double x, out double y, out double width, out double height)) {
                return TryCharge(_viewportPixels);
            }
            if (width < 0D || height < 0D) return false;

            (double left, double top, double right, double bottom) = transform.TransformRectangleBounds(x, y, width, height);
            left = Math.Max(left, _viewLeft);
            top = Math.Max(top, _viewTop);
            right = Math.Min(right, _viewRight);
            bottom = Math.Min(bottom, _viewBottom);
            if (right < left || bottom < top) return true;
            double pixels = (right - left) * _pixelScaleX * (bottom - top) * _pixelScaleY;
            return TryCharge(Math.Max(1D, pixels));
        }

        internal bool TryChargeFilterDefinition(XElement definition) {
            if (!definition.Name.LocalName.Equals("filter", StringComparison.OrdinalIgnoreCase)) return true;
            foreach (XElement primitive in definition.Descendants()) {
                string name = primitive.Name.LocalName.ToLowerInvariant();
                double multiplier;
                switch (name) {
                    case "fegaussianblur":
                        if (!TryReadPair(primitive, "stdDeviation", 0D, out double sigmaX, out double sigmaY)) return false;
                        multiplier = 1D + 6D * (sigmaX + sigmaY);
                        break;
                    case "femorphology":
                        if (!TryReadPair(primitive, "radius", 0D, out double radiusX, out double radiusY)) return false;
                        multiplier = 1D + 2D * (radiusX + radiusY);
                        break;
                    case "feconvolvematrix":
                        if (!TryReadPair(primitive, "order", 3D, out double orderX, out double orderY)
                            || orderX != Math.Truncate(orderX)
                            || orderY != Math.Truncate(orderY)
                            || orderX < 1D
                            || orderY < 1D) return false;
                        multiplier = orderX * orderY;
                        break;
                    case "feturbulence":
                        string? octaveText = ReadRasterProjectedAttribute(primitive, "numOctaves");
                        if (string.IsNullOrWhiteSpace(octaveText)) {
                            multiplier = 1D;
                        } else if (!double.TryParse(octaveText, System.Globalization.NumberStyles.Float,
                                       System.Globalization.CultureInfo.InvariantCulture, out double octaves)
                                   || octaves != Math.Truncate(octaves)
                                   || octaves < 1D) {
                            return false;
                        } else {
                            multiplier = octaves;
                        }
                        break;
                    default:
                        continue;
                }
                if (!TryCharge(_viewportPixels * multiplier)) return false;
            }
            return true;
        }

        private bool TryCharge(double work) {
            if (double.IsNaN(work) || double.IsInfinity(work) || work < 0D || work > _remainingWork) return false;
            _remainingWork -= work;
            return true;
        }

        private static bool TryGetPaintBounds(
            XElement element,
            string name,
            out double x,
            out double y,
            out double width,
            out double height) {
            x = y = width = height = 0D;
            switch (name) {
                case "rect":
                case "image":
                case "foreignobject":
                    return TryReadLength(element, "x", 0D, out x)
                        && TryReadLength(element, "y", 0D, out y)
                        && TryReadLength(element, "width", 0D, out width)
                        && TryReadLength(element, "height", 0D, out height);
                case "circle":
                    if (!TryReadLength(element, "cx", 0D, out double cx)
                        || !TryReadLength(element, "cy", 0D, out double cy)
                        || !TryReadLength(element, "r", 0D, out double radius)) return false;
                    x = cx - radius;
                    y = cy - radius;
                    width = height = radius * 2D;
                    return true;
                case "ellipse":
                    if (!TryReadLength(element, "cx", 0D, out cx)
                        || !TryReadLength(element, "cy", 0D, out cy)
                        || !TryReadLength(element, "rx", 0D, out double radiusX)
                        || !TryReadLength(element, "ry", 0D, out double radiusY)) return false;
                    x = cx - radiusX;
                    y = cy - radiusY;
                    width = radiusX * 2D;
                    height = radiusY * 2D;
                    return true;
                case "line":
                    if (!TryReadLength(element, "x1", 0D, out double x1)
                        || !TryReadLength(element, "y1", 0D, out double y1)
                        || !TryReadLength(element, "x2", 0D, out double x2)
                        || !TryReadLength(element, "y2", 0D, out double y2)) return false;
                    return SetBounds(x1, y1, x2, y2, out x, out y, out width, out height);
                case "polygon":
                case "polyline":
                    if (!TryParseNumberList(ReadRasterProjectedAttribute(element, "points"), MaximumSvgPathCommands * 2,
                            out IReadOnlyList<double> values, out bool valueLimitExceeded)
                        || valueLimitExceeded
                        || values.Count < 2) return false;
                    return TrySetPointBounds(values, out x, out y, out width, out height);
                case "path":
                    if (!OfficeSvgPathDataParser.TryParse(ReadRasterProjectedAttribute(element, "d"), MaximumSvgPathCommands,
                            out IReadOnlyList<OfficePathCommand> commands, out bool commandLimitExceeded)
                        || commandLimitExceeded) return false;
                    return TrySetPathBounds(commands, out x, out y, out width, out height);
                case "text":
                case "tspan":
                case "textpath":
                    if (!TryReadLength(element, "x", 0D, out x)
                        || !TryReadLength(element, "y", 0D, out y)) return false;
                    double fontSize = 16D;
                    string? fontSizeText = ReadRasterPresentationProperty(element, "font-size");
                    if (!string.IsNullOrWhiteSpace(fontSizeText)
                        && !OfficeImageReader.TryParseSvgLength(fontSizeText, out fontSize)) return false;
                    width = Math.Max(1D, element.Value.Length * fontSize);
                    height = Math.Max(1D, fontSize * 1.5D);
                    y -= fontSize;
                    return true;
                default:
                    return false;
            }
        }

        private static bool TryReadLength(XElement element, string name, double defaultValue, out double value) {
            string? text = ReadRasterProjectedAttribute(element, name);
            if (string.IsNullOrWhiteSpace(text)) {
                value = defaultValue;
                return true;
            }
            return OfficeImageReader.TryParseSvgLength(text, out value)
                && !double.IsNaN(value)
                && !double.IsInfinity(value);
        }

        private static bool TryReadPair(
            XElement element,
            string name,
            double defaultValue,
            out double first,
            out double second) {
            string? text = ReadRasterProjectedAttribute(element, name);
            if (string.IsNullOrWhiteSpace(text)) {
                first = second = defaultValue;
                return true;
            }
            if (!TryParseNumberList(text, 2, out IReadOnlyList<double> values, out bool exceeded)
                || exceeded
                || values.Count is < 1 or > 2
                || values[0] < 0D
                || (values.Count == 2 && values[1] < 0D)) {
                first = second = 0D;
                return false;
            }
            first = values[0];
            second = values.Count == 2 ? values[1] : first;
            return true;
        }

        private static bool TrySetPointBounds(
            IReadOnlyList<double> values,
            out double x,
            out double y,
            out double width,
            out double height) {
            double minX = double.PositiveInfinity;
            double minY = double.PositiveInfinity;
            double maxX = double.NegativeInfinity;
            double maxY = double.NegativeInfinity;
            for (int index = 0; index + 1 < values.Count; index += 2) {
                AddPoint(values[index], values[index + 1], ref minX, ref minY, ref maxX, ref maxY);
            }
            return SetBounds(minX, minY, maxX, maxY, out x, out y, out width, out height);
        }

        private static bool TrySetPathBounds(
            IReadOnlyList<OfficePathCommand> commands,
            out double x,
            out double y,
            out double width,
            out double height) {
            double minX = double.PositiveInfinity;
            double minY = double.PositiveInfinity;
            double maxX = double.NegativeInfinity;
            double maxY = double.NegativeInfinity;
            foreach (OfficePathCommand command in commands) {
                if (command.Kind == OfficePathCommandKind.Close) continue;
                AddPoint(command.Point.X, command.Point.Y, ref minX, ref minY, ref maxX, ref maxY);
                if (command.Kind is OfficePathCommandKind.QuadraticBezierTo or OfficePathCommandKind.CubicBezierTo) {
                    AddPoint(command.ControlPoint1.X, command.ControlPoint1.Y, ref minX, ref minY, ref maxX, ref maxY);
                }
                if (command.Kind == OfficePathCommandKind.CubicBezierTo) {
                    AddPoint(command.ControlPoint2.X, command.ControlPoint2.Y, ref minX, ref minY, ref maxX, ref maxY);
                }
            }
            return SetBounds(minX, minY, maxX, maxY, out x, out y, out width, out height);
        }

        private static void AddPoint(
            double pointX,
            double pointY,
            ref double minX,
            ref double minY,
            ref double maxX,
            ref double maxY) {
            minX = Math.Min(minX, pointX);
            minY = Math.Min(minY, pointY);
            maxX = Math.Max(maxX, pointX);
            maxY = Math.Max(maxY, pointY);
        }

        private static bool SetBounds(
            double minX,
            double minY,
            double maxX,
            double maxY,
            out double x,
            out double y,
            out double width,
            out double height) {
            x = minX;
            y = minY;
            width = maxX - minX;
            height = maxY - minY;
            return !(double.IsNaN(width) || double.IsInfinity(width) || double.IsNaN(height) || double.IsInfinity(height));
        }
    }
}
