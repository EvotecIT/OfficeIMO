using System.Collections.Generic;
using System.Globalization;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Drawing;
using W = DocumentFormat.OpenXml.Wordprocessing;
using V = DocumentFormat.OpenXml.Vml;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Word.Pdf {
    public static partial class WordPdfConverterExtensions {
        private static bool TryGetNativeVmlBox(OpenXmlElement element, NativeVmlFrame frame, double pageWidth, double pageHeight, out NativeVmlBox box) {
            Dictionary<string, string> style = ParseNativeVmlStyle(GetNativeOpenXmlAttribute(element, "style"));
            double width = ResolveNativeVmlLength(style.TryGetValue("width", out string? widthValue) ? widthValue : null, frame.Width, frame.CoordWidth) ??
                           ResolveNativeVmlPercent(style, "mso-width-percent", pageWidth) ?? 0D;
            double height = ResolveNativeVmlLength(style.TryGetValue("height", out string? heightValue) ? heightValue : null, frame.Height, frame.CoordHeight) ??
                            ResolveNativeVmlPercent(style, "mso-height-percent", pageHeight) ?? 0D;

            double? xPercent = ResolveNativeVmlPercent(style, "mso-left-percent", pageWidth);
            double? yPercent = ResolveNativeVmlPercent(style, "mso-top-percent", pageHeight);
            bool hasHorizontalPosition = style.ContainsKey("mso-position-horizontal");
            bool hasVerticalPosition = style.ContainsKey("mso-position-vertical");
            bool hasExplicitX = xPercent.HasValue ||
                                HasNativeVmlExplicitPosition(style, "left", hasHorizontalPosition) ||
                                HasNativeVmlExplicitPosition(style, "margin-left", hasHorizontalPosition);
            bool hasExplicitY = yPercent.HasValue ||
                                HasNativeVmlExplicitPosition(style, "top", hasVerticalPosition) ||
                                HasNativeVmlExplicitPosition(style, "margin-top", hasVerticalPosition);
            double x = xPercent ??
                       ResolveNativeVmlPosition(style.TryGetValue("left", out string? left) ? left : null, frame.Width, frame.CoordWidth, frame.CoordOriginX) ??
                       ResolveNativeVmlPosition(style.TryGetValue("margin-left", out string? marginLeft) ? marginLeft : null, frame.Width, frame.CoordWidth, frame.CoordOriginX) ?? 0D;
            double y = yPercent ??
                       ResolveNativeVmlPosition(style.TryGetValue("top", out string? top) ? top : null, frame.Height, frame.CoordHeight, frame.CoordOriginY) ??
                       ResolveNativeVmlPosition(style.TryGetValue("margin-top", out string? marginTop) ? marginTop : null, frame.Height, frame.CoordHeight, frame.CoordOriginY) ?? 0D;

            if (element.LocalName.Equals("line", StringComparison.OrdinalIgnoreCase) &&
                (width <= 0D || height <= 0D) &&
                TryGetNativeVmlLineBounds(element, frame, out double lineX, out double lineY, out double lineWidth, out double lineHeight)) {
                x += lineX;
                y += lineY;
                width = Math.Max(width, lineWidth);
                height = Math.Max(height, lineHeight);
            }

            if (style.TryGetValue("mso-position-horizontal", out string? horizontalPosition) && !hasExplicitX) {
                if (horizontalPosition.Equals("center", StringComparison.OrdinalIgnoreCase)) {
                    x = (pageWidth - width) / 2D;
                } else if (horizontalPosition.Equals("right", StringComparison.OrdinalIgnoreCase)) {
                    x = pageWidth - width;
                }
            }

            if (style.TryGetValue("mso-position-vertical", out string? verticalPosition) &&
                !hasExplicitY) {
                if (verticalPosition.Equals("center", StringComparison.OrdinalIgnoreCase)) {
                    y = (pageHeight - height) / 2D;
                } else if (verticalPosition.Equals("bottom", StringComparison.OrdinalIgnoreCase)) {
                    y = pageHeight - height;
                }
            }

            box = new NativeVmlBox(frame.X + x, frame.Y + y, width, height);
            return width > 0D && height > 0D;
        }

        private static bool TryGetNativeVmlLineBounds(OpenXmlElement element, NativeVmlFrame frame, out double x, out double y, out double width, out double height) {
            (double x1, double y1) = ParseNativeShapePoint(GetNativeOpenXmlAttribute(element, "from") ?? "0pt,0pt", frame);
            (double x2, double y2) = ParseNativeShapePoint(GetNativeOpenXmlAttribute(element, "to") ?? "0pt,0pt", frame);
            x = Math.Min(x1, x2);
            y = Math.Min(y1, y2);
            width = Math.Abs(x2 - x1);
            height = Math.Abs(y2 - y1);
            if (width <= 0D && height <= 0D) {
                return false;
            }

            width = Math.Max(width, 0.01D);
            height = Math.Max(height, 0.01D);
            return true;
        }

        private static (double X, double Y) ParseNativeShapePoint(string value, NativeVmlFrame frame) {
            string[] parts = value.Split(',');
            if (parts.Length != 2) {
                return (0D, 0D);
            }

            double x = ResolveNativeVmlPosition(parts[0], frame.Width, frame.CoordWidth, frame.CoordOriginX) ?? 0D;
            double y = ResolveNativeVmlPosition(parts[1], frame.Height, frame.CoordHeight, frame.CoordOriginY) ?? 0D;
            return (x, y);
        }

        private static bool HasNativeVmlExplicitPosition(Dictionary<string, string> style, string key, bool hasRelativePosition) {
            if (!style.TryGetValue(key, out string? value)) {
                return false;
            }

            if (hasRelativePosition) {
                double? resolved = ResolveNativeVmlPosition(value, 1D, 1D, 0D);
                if (resolved.HasValue && Math.Abs(resolved.Value) < 0.001D) {
                    return false;
                }
            }

            return true;
        }

        private static Dictionary<string, string> ParseNativeVmlStyle(string? style) {
            var values = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            if (string.IsNullOrWhiteSpace(style)) {
                return values;
            }

            foreach (string part in style!.Split(';')) {
                int separator = part.IndexOf(':');
                if (separator <= 0 || separator == part.Length - 1) {
                    continue;
                }

                values[part.Substring(0, separator).Trim()] = part.Substring(separator + 1).Trim();
            }

            return values;
        }

        private static double? ResolveNativeVmlPosition(string? value, double parentSize, double parentCoord, double parentCoordOrigin) {
            if (string.IsNullOrWhiteSpace(value)) {
                return null;
            }

            string normalized = value!.Trim();
            if (HasNativeVmlLengthUnit(normalized) || normalized.EndsWith("%", StringComparison.OrdinalIgnoreCase)) {
                return ResolveNativeVmlLength(normalized, parentSize, parentCoord);
            }

            double? number = ParseNativeVmlDouble(normalized);
            if (!number.HasValue) {
                return null;
            }

            double relative = number.Value - parentCoordOrigin;
            return parentCoord > 0D && Math.Abs(parentCoord - parentSize) > 0.01D
                ? relative / parentCoord * parentSize
                : relative;
        }

        private static double? ResolveNativeVmlLength(string? value, double parentSize, double parentCoord) {
            if (string.IsNullOrWhiteSpace(value)) {
                return null;
            }

            string normalized = value!.Trim();
            if (normalized.EndsWith("pt", StringComparison.OrdinalIgnoreCase)) return NormalizeNativeVmlLength(ParseNativeVmlDouble(normalized.Substring(0, normalized.Length - 2)));
            if (normalized.EndsWith("in", StringComparison.OrdinalIgnoreCase)) return MultiplyNativeVmlLength(ParseNativeVmlDouble(normalized.Substring(0, normalized.Length - 2)), 72D);
            if (normalized.EndsWith("cm", StringComparison.OrdinalIgnoreCase)) return MultiplyNativeVmlLength(ParseNativeVmlDouble(normalized.Substring(0, normalized.Length - 2)), 28.3464566929D);
            if (normalized.EndsWith("mm", StringComparison.OrdinalIgnoreCase)) return MultiplyNativeVmlLength(ParseNativeVmlDouble(normalized.Substring(0, normalized.Length - 2)), 2.83464566929D);
            if (normalized.EndsWith("px", StringComparison.OrdinalIgnoreCase)) return MultiplyNativeVmlLength(ParseNativeVmlDouble(normalized.Substring(0, normalized.Length - 2)), 0.75D);
            if (normalized.EndsWith("%", StringComparison.OrdinalIgnoreCase)) return MultiplyNativeVmlLength(ParseNativeVmlDouble(normalized.Substring(0, normalized.Length - 1)), parentSize / 100D);

            double? number = ParseNativeVmlDouble(normalized);
            if (!number.HasValue) {
                return null;
            }

            double resolved = parentCoord > 0D && Math.Abs(parentCoord - parentSize) > 0.01D
                ? number.Value / parentCoord * parentSize
                : number.Value;
            return NormalizeNativeVmlLength(resolved);
        }

        private static bool HasNativeVmlLengthUnit(string value) =>
            value.EndsWith("pt", StringComparison.OrdinalIgnoreCase) ||
            value.EndsWith("in", StringComparison.OrdinalIgnoreCase) ||
            value.EndsWith("cm", StringComparison.OrdinalIgnoreCase) ||
            value.EndsWith("mm", StringComparison.OrdinalIgnoreCase) ||
            value.EndsWith("px", StringComparison.OrdinalIgnoreCase);

        private static double? ResolveNativeVmlPercent(Dictionary<string, string> style, string key, double reference) {
            if (!style.TryGetValue(key, out string? value) ||
                !double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out double percent) ||
                percent <= 0D) {
                return null;
            }

            return reference * percent / 1000D;
        }

        private static (double Width, double Height) GetNativeVmlCoordSize(OpenXmlElement element, double fallbackWidth, double fallbackHeight) {
            string? coordSize = GetNativeOpenXmlAttribute(element, "coordsize");
            if (string.IsNullOrWhiteSpace(coordSize)) {
                return (fallbackWidth, fallbackHeight);
            }

            string[] parts = coordSize!.Split(',');
            if (parts.Length == 2 &&
                double.TryParse(parts[0], NumberStyles.Float, CultureInfo.InvariantCulture, out double width) &&
                double.TryParse(parts[1], NumberStyles.Float, CultureInfo.InvariantCulture, out double height) &&
                width > 0D &&
                height > 0D) {
                return (width, height);
            }

            return (fallbackWidth, fallbackHeight);
        }

        private static (double X, double Y) GetNativeVmlCoordOrigin(OpenXmlElement element) {
            string? coordOrigin = GetNativeOpenXmlAttribute(element, "coordorigin");
            if (string.IsNullOrWhiteSpace(coordOrigin)) {
                return (0D, 0D);
            }

            string[] parts = coordOrigin!.Split(',');
            if (parts.Length == 2 &&
                double.TryParse(parts[0], NumberStyles.Float, CultureInfo.InvariantCulture, out double x) &&
                double.TryParse(parts[1], NumberStyles.Float, CultureInfo.InvariantCulture, out double y)) {
                return (x, y);
            }

            return (0D, 0D);
        }

        private static string? NormalizeNativeVmlColor(string? value) {
            if (string.IsNullOrWhiteSpace(value)) {
                return null;
            }

            string trimmed = value!.Trim();
            int space = trimmed.IndexOf(' ');
            if (space > 0) {
                trimmed = trimmed.Substring(0, space);
            }

            return trimmed.Equals("none", StringComparison.OrdinalIgnoreCase) ? null : trimmed;
        }

        private static bool IsNativeVmlNoColor(string? value) =>
            value?.Trim().Equals("none", StringComparison.OrdinalIgnoreCase) == true;

        private static string? GetNativeOpenXmlAttribute(OpenXmlElement element, string localName) {
            foreach (OpenXmlAttribute attribute in element.GetAttributes()) {
                if (attribute.LocalName.Equals(localName, StringComparison.OrdinalIgnoreCase)) {
                    return attribute.Value;
                }
            }

            return null;
        }

        private static bool IsNativeVmlSwitchEnabled(string? value) =>
            string.IsNullOrWhiteSpace(value) ||
            (!value!.Trim().Equals("f", StringComparison.OrdinalIgnoreCase) &&
             !value.Trim().Equals("false", StringComparison.OrdinalIgnoreCase) &&
             !value.Trim().Equals("0", StringComparison.OrdinalIgnoreCase));

        private static bool IsNativeVmlHidden(OpenXmlElement element) {
            Dictionary<string, string> style = ParseNativeVmlStyle(GetNativeOpenXmlAttribute(element, "style"));
            return style.TryGetValue("visibility", out string? value) &&
                   value.Equals("hidden", StringComparison.OrdinalIgnoreCase);
        }

        private static double? ParseNativeVmlStrokeWeight(string? value) {
            if (string.IsNullOrWhiteSpace(value)) return null;
            double? unitless = ParseNativeVmlDouble(value!.Trim());
            // Unitless VML stroke weights use EMUs, unlike shape coordinates.
            return unitless.HasValue
                ? NormalizeNativeVmlLength(unitless.Value / 12700D)
                : ResolveNativeVmlLength(value, 1D, 1D);
        }

        private static double GetNativeVmlFirstAdjustment(OpenXmlElement element, double fallback) {
            string? value = GetNativeOpenXmlAttribute(element, "adj");
            if (string.IsNullOrWhiteSpace(value)) {
                return fallback;
            }

            string[] parts = value!.Split(new[] { ',', ';', ' ' }, StringSplitOptions.RemoveEmptyEntries);
            double? adjustment = parts.Length > 0 ? ParseNativeVmlDouble(parts[0]) : null;
            return adjustment.HasValue
                ? adjustment.Value
                : fallback;
        }

        private static double GetNativeVmlRoundRectCornerRadius(OpenXmlElement element, double width, double height) {
            const double defaultArcSize = 0.2D;
            double fraction = ParseNativeVmlOpacity(GetNativeOpenXmlAttribute(element, "arcsize")) ?? defaultArcSize;
            fraction = Math.Max(0D, Math.Min(0.5D, fraction));
            return Math.Min(width, height) * fraction;
        }

        private static double? ParseNativeVmlDouble(string value) {
            string normalized = value.Trim();
            if (double.TryParse(normalized, NumberStyles.Float, CultureInfo.InvariantCulture, out double result) && IsNativeVmlFinite(result)) {
                return result;
            }

            if (double.TryParse(normalized.Replace(',', '.'), NumberStyles.Float, CultureInfo.InvariantCulture, out result) && IsNativeVmlFinite(result)) {
                return result;
            }

            return null;
        }

        private static double? MultiplyNativeVmlLength(double? value, double factor) =>
            value.HasValue ? NormalizeNativeVmlLength(value.Value * factor) : null;

        private static double? NormalizeNativeVmlLength(double? value) {
            if (!value.HasValue || !IsNativeVmlFinite(value.Value) || Math.Abs(value.Value) > MaxNativeVmlLengthPoints) {
                return null;
            }

            return value.Value;
        }

        private static bool IsNativeVmlFinite(double value) =>
            !double.IsNaN(value) && !double.IsInfinity(value);

        private static double? ParseNativeVmlOpacity(string? value) {
            if (string.IsNullOrWhiteSpace(value)) {
                return null;
            }

            string normalized = value!.Trim();
            double? parsed;
            if (normalized.EndsWith("%", StringComparison.OrdinalIgnoreCase)) {
                parsed = ParseNativeVmlDouble(normalized.Substring(0, normalized.Length - 1));
                if (parsed.HasValue) {
                    parsed /= 100D;
                }
            } else if (normalized.EndsWith("f", StringComparison.OrdinalIgnoreCase)) {
                parsed = ParseNativeVmlDouble(normalized.Substring(0, normalized.Length - 1));
                if (parsed.HasValue) {
                    parsed /= 65536D;
                }
            } else {
                parsed = ParseNativeVmlDouble(normalized);
            }

            if (!parsed.HasValue) {
                return null;
            }

            return Math.Max(0D, Math.Min(1D, parsed.Value));
        }

        private static bool IsNativeVmlBoxInsidePage(NativeVmlBox box, double pageWidth, double pageHeight) =>
            box.X >= 0D &&
            box.Y >= 0D &&
            box.Width > 0D &&
            box.Height > 0D &&
            box.X + box.Width <= pageWidth + 0.01D &&
            box.Y + box.Height <= pageHeight + 0.01D;

        private static bool IsNativeVmlBoxVisibleOnPage(NativeVmlBox box, double pageWidth, double pageHeight) =>
            box.Width > 0D &&
            box.Height > 0D &&
            box.X < pageWidth &&
            box.Y < pageHeight &&
            box.X + box.Width > 0D &&
            box.Y + box.Height > 0D;

        private static void RenderNativeVmlVisible(PdfCore.PdfPageCanvas canvas, NativeVmlBox box, double pageWidth, double pageHeight, Action<PdfCore.PdfPageCanvas> render) {
            if (IsNativeVmlBoxInsidePage(box, pageWidth, pageHeight)) {
                render(canvas);
                return;
            }

            canvas.Clip(0D, 0D, pageWidth, pageHeight, render);
        }

        private readonly struct NativeVmlFrame {
            public NativeVmlFrame(double x, double y, double width, double height, double coordWidth, double coordHeight, double coordOriginX, double coordOriginY) {
                X = x;
                Y = y;
                Width = width;
                Height = height;
                CoordWidth = coordWidth;
                CoordHeight = coordHeight;
                CoordOriginX = coordOriginX;
                CoordOriginY = coordOriginY;
            }

            public double X { get; }
            public double Y { get; }
            public double Width { get; }
            public double Height { get; }
            public double CoordWidth { get; }
            public double CoordHeight { get; }
            public double CoordOriginX { get; }
            public double CoordOriginY { get; }
        }

        private readonly struct NativeVmlBox {
            public NativeVmlBox(double x, double y, double width, double height) {
                X = x;
                Y = y;
                Width = width;
                Height = height;
            }

            public double X { get; }
            public double Y { get; }
            public double Width { get; }
            public double Height { get; }
        }
    }
}
