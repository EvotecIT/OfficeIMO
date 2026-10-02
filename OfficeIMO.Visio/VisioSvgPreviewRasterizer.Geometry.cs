using System;
using System.Collections.Generic;
using System.Globalization;
using System.Xml.Linq;
using OfficeIMO.Drawing;

namespace OfficeIMO.Visio {
    internal static partial class VisioSvgPreviewRasterizer {
        private static bool TryReadViewBoxTransform(XElement definition, XElement useElement, out SvgTransform transform) {
            transform = SvgTransform.Identity;
            if (!TryParseNumbers(definition.Attribute("viewBox")?.Value, out List<double> viewBox) ||
                viewBox.Count < 4 ||
                viewBox[2] <= 0D ||
                viewBox[3] <= 0D) {
                return false;
            }

            double width = ReadLength(useElement, "width", viewBox[2]);
            double height = ReadLength(useElement, "height", viewBox[3]);
            if (width <= 0D || height <= 0D) {
                return false;
            }

            transform = CreateViewBoxTransform(viewBox[0], viewBox[1], viewBox[2], viewBox[3], 0D, 0D, width, height, useElement.Attribute("preserveAspectRatio")?.Value ?? definition.Attribute("preserveAspectRatio")?.Value);
            return true;
        }

        private static SvgTransform CreateViewBoxTransform(double viewLeft, double viewTop, double viewWidth, double viewHeight, double viewportX, double viewportY, double viewportWidth, double viewportHeight, string? preserveAspectRatio) {
            string align = "xMidYMid";
            string meetOrSlice = "meet";
            if (!string.IsNullOrWhiteSpace(preserveAspectRatio)) {
                string[] parts = preserveAspectRatio!.Split(new[] { ' ', '\t', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
                int partOffset = parts.Length > 0 && string.Equals(parts[0], "defer", StringComparison.OrdinalIgnoreCase) ? 1 : 0;

                if (parts.Length > partOffset) {
                    align = parts[partOffset];
                }

                if (parts.Length > partOffset + 1) {
                    meetOrSlice = parts[partOffset + 1];
                }
            }

            double scaleX = viewportWidth / viewWidth;
            double scaleY = viewportHeight / viewHeight;
            if (string.Equals(align, "none", StringComparison.OrdinalIgnoreCase)) {
                return SvgTransform.Create(scaleX, 0D, 0D, scaleY, viewportX - (viewLeft * scaleX), viewportY - (viewTop * scaleY));
            }

            double scale = string.Equals(meetOrSlice, "slice", StringComparison.OrdinalIgnoreCase)
                ? Math.Max(scaleX, scaleY)
                : Math.Min(scaleX, scaleY);
            double renderedWidth = viewWidth * scale;
            double renderedHeight = viewHeight * scale;
            double offsetX = align.IndexOf("xMax", StringComparison.OrdinalIgnoreCase) >= 0
                ? viewportWidth - renderedWidth
                : align.IndexOf("xMid", StringComparison.OrdinalIgnoreCase) >= 0
                    ? (viewportWidth - renderedWidth) / 2D
                    : 0D;
            double offsetY = align.IndexOf("YMax", StringComparison.OrdinalIgnoreCase) >= 0
                ? viewportHeight - renderedHeight
                : align.IndexOf("YMid", StringComparison.OrdinalIgnoreCase) >= 0
                    ? (viewportHeight - renderedHeight) / 2D
                    : 0D;

            return SvgTransform.Create(scale, 0D, 0D, scale, viewportX + offsetX - (viewLeft * scale), viewportY + offsetY - (viewTop * scale));
        }

        private static void ResolveViewport(XElement root, out double viewLeft, out double viewTop, out double viewWidth, out double viewHeight, out int width, out int height) {
            viewLeft = 0D;
            viewTop = 0D;
            viewWidth = ReadLength(root, "width", DefaultSize);
            viewHeight = ReadLength(root, "height", DefaultSize);
            if (TryParseNumbers(root.Attribute("viewBox")?.Value, out List<double> viewBox) && viewBox.Count >= 4 && viewBox[2] > 0D && viewBox[3] > 0D) {
                viewLeft = viewBox[0];
                viewTop = viewBox[1];
                viewWidth = viewBox[2];
                viewHeight = viewBox[3];
            }

            double rawWidth = ReadLength(root, "width", viewWidth);
            double rawHeight = ReadLength(root, "height", viewHeight);
            if (rawWidth <= 0D) {
                rawWidth = viewWidth;
            }

            if (rawHeight <= 0D) {
                rawHeight = viewHeight;
            }

            width = ClampSize((int)Math.Round(rawWidth));
            height = ClampSize((int)Math.Round(rawHeight));
        }

        private static int ClampSize(int value) => Math.Max(1, Math.Min(MaximumSize, value));

        private static double ReadLength(XElement element, string name, double fallback) =>
            TryParseLength(element.Attribute(name)?.Value, out double value) ? value : fallback;

        private static double ReadLength(XElement element, string name, double fallback, SvgRenderContext context, SvgLengthAxis axis) =>
            ReadLength(element, name, fallback, GetLengthReference(context, axis));

        private static double ReadLength(XElement element, string name, double fallback, double? percentageReference) =>
            TryParseLength(element.Attribute(name)?.Value, percentageReference, out double value) ? value : fallback;

        private static double? GetLengthReference(SvgRenderContext context, SvgLengthAxis axis) {
            SvgPaintBounds viewport = context.ViewportBounds;
            double width = viewport.Width;
            double height = viewport.Height;
            if (width <= 0D || height <= 0D) {
                return null;
            }

            return axis switch {
                SvgLengthAxis.X => width,
                SvgLengthAxis.Y => height,
                _ => Math.Sqrt((width * width) + (height * height)) / Math.Sqrt(2D)
            };
        }

        private static SvgTransform ReadTransform(string? value) {
            if (string.IsNullOrWhiteSpace(value)) {
                return SvgTransform.Identity;
            }

            SvgTransform transform = SvgTransform.Identity;
            int offset = 0;
            while (offset < value!.Length) {
                while (offset < value.Length && char.IsWhiteSpace(value[offset])) {
                    offset++;
                }

                int nameStart = offset;
                while (offset < value.Length && char.IsLetter(value[offset])) {
                    offset++;
                }

                string name = value.Substring(nameStart, offset - nameStart);
                if (string.IsNullOrEmpty(name) || offset >= value.Length || value[offset] != '(') {
                    break;
                }

                int close = value.IndexOf(')', offset + 1);
                if (close < 0) {
                    break;
                }

                if (TryParseNumbers(value.Substring(offset + 1, close - offset - 1), out List<double> numbers)) {
                    if (string.Equals(name, "translate", StringComparison.OrdinalIgnoreCase) && numbers.Count >= 1) {
                        transform = transform.Multiply(SvgTransform.Create(1D, 0D, 0D, 1D, numbers[0], numbers.Count > 1 ? numbers[1] : 0D));
                    } else if (string.Equals(name, "scale", StringComparison.OrdinalIgnoreCase) && numbers.Count >= 1) {
                        transform = transform.Multiply(SvgTransform.Create(numbers[0], 0D, 0D, numbers.Count > 1 ? numbers[1] : numbers[0], 0D, 0D));
                    } else if (string.Equals(name, "rotate", StringComparison.OrdinalIgnoreCase) && numbers.Count >= 1) {
                        transform = transform.Multiply(CreateRotationTransform(numbers));
                    } else if (string.Equals(name, "skewX", StringComparison.OrdinalIgnoreCase) && numbers.Count >= 1) {
                        transform = transform.Multiply(SvgTransform.Create(1D, 0D, Math.Tan(OfficeGeometry.DegreesToRadians(numbers[0])), 1D, 0D, 0D));
                    } else if (string.Equals(name, "skewY", StringComparison.OrdinalIgnoreCase) && numbers.Count >= 1) {
                        transform = transform.Multiply(SvgTransform.Create(1D, Math.Tan(OfficeGeometry.DegreesToRadians(numbers[0])), 0D, 1D, 0D, 0D));
                    } else if (string.Equals(name, "matrix", StringComparison.OrdinalIgnoreCase) && numbers.Count >= 6) {
                        transform = transform.Multiply(SvgTransform.Create(numbers[0], numbers[1], numbers[2], numbers[3], numbers[4], numbers[5]));
                    }
                }

                offset = close + 1;
                while (offset < value.Length && (char.IsWhiteSpace(value[offset]) || value[offset] == ',')) {
                    offset++;
                }
            }

            return transform;
        }

        private static SvgTransform CreateRotationTransform(IReadOnlyList<double> numbers) {
            double radians = OfficeGeometry.DegreesToRadians(numbers[0]);
            double cos = Math.Cos(radians);
            double sin = Math.Sin(radians);
            SvgTransform rotate = SvgTransform.Create(cos, sin, -sin, cos, 0D, 0D);
            if (numbers.Count < 3) {
                return rotate;
            }

            double cx = numbers[1];
            double cy = numbers[2];
            return SvgTransform.Create(1D, 0D, 0D, 1D, cx, cy)
                .Multiply(rotate)
                .Multiply(SvgTransform.Create(1D, 0D, 0D, 1D, -cx, -cy));
        }

        private static bool TryParseLength(string? value, out double result) {
            return TryParseLength(value, null, out result);
        }

        private static bool TryParseLength(string? value, double? percentageReference, out double result) {
            result = 0D;
            if (string.IsNullOrWhiteSpace(value)) {
                return false;
            }

            string trimmed = value!.Trim();
            if (trimmed.EndsWith("%", StringComparison.Ordinal)) {
                if (!percentageReference.HasValue || percentageReference.Value <= 0D) {
                    return false;
                }

                string rawPercentage = trimmed.Substring(0, trimmed.Length - 1);
                if (!double.TryParse(rawPercentage, NumberStyles.Float, CultureInfo.InvariantCulture, out double percentage)) {
                    return false;
                }

                result = percentageReference.Value * percentage / 100D;
                return true;
            }

            int end = 0;
            while (end < trimmed.Length && (char.IsDigit(trimmed[end]) || trimmed[end] == '-' || trimmed[end] == '+' || trimmed[end] == '.' || trimmed[end] == 'e' || trimmed[end] == 'E')) {
                end++;
            }

            if (end == 0 || !double.TryParse(trimmed.Substring(0, end), NumberStyles.Float, CultureInfo.InvariantCulture, out result)) {
                return false;
            }

            string unit = trimmed.Substring(end).Trim();
            switch (unit.ToLowerInvariant()) {
                case "":
                case "px":
                    return true;
                case "in":
                    result *= 96D;
                    return true;
                case "cm":
                    result *= 96D / 2.54D;
                    return true;
                case "mm":
                    result *= 96D / 25.4D;
                    return true;
                case "q":
                    result *= 96D / 101.6D;
                    return true;
                case "pt":
                    result *= 96D / 72D;
                    return true;
                case "pc":
                    result *= 16D;
                    return true;
                default:
                    result = 0D;
                    return false;
            }
        }

        private enum SvgLengthAxis {
            X,
            Y,
            Diagonal
        }

        private static bool TryParsePoints(string? value, out List<(double X, double Y)> points) {
            points = new List<(double X, double Y)>();
            if (!TryParseNumbers(value, out List<double> numbers) || numbers.Count < 4) {
                return false;
            }

            for (int i = 0; i + 1 < numbers.Count; i += 2) {
                points.Add((numbers[i], numbers[i + 1]));
            }

            return points.Count >= 2;
        }

        private static bool TryParseNumbers(string? value, out List<double> numbers) {
            numbers = new List<double>();
            if (string.IsNullOrWhiteSpace(value)) {
                return false;
            }

            int index = 0;
            while (index < value!.Length) {
                while (index < value.Length && (char.IsWhiteSpace(value[index]) || value[index] == ',')) {
                    index++;
                }

                int start = index;
                if (index < value.Length && (value[index] == '-' || value[index] == '+')) {
                    index++;
                }

                while (index < value.Length && (char.IsDigit(value[index]) || value[index] == '.')) {
                    index++;
                }

                if (index < value.Length && (value[index] == 'e' || value[index] == 'E')) {
                    index++;
                    if (index < value.Length && (value[index] == '-' || value[index] == '+')) {
                        index++;
                    }

                    while (index < value.Length && char.IsDigit(value[index])) {
                        index++;
                    }
                }

                if (index == start) {
                    break;
                }

                if (!double.TryParse(value.Substring(start, index - start), NumberStyles.Float, CultureInfo.InvariantCulture, out double number)) {
                    return false;
                }

                numbers.Add(number);
            }

            return numbers.Count > 0;
        }

    }
}
