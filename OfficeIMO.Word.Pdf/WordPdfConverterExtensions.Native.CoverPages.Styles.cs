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
        private static void ApplyNativeVmlShapeStyle(OfficeShape shape, OpenXmlElement element) {
            if (shape.Kind != OfficeShapeKind.Line) {
                OpenXmlElement? fillElement = GetNativeVmlChild(element, "fill");
                bool fillEnabled = IsNativeVmlSwitchEnabled(GetNativeOpenXmlAttribute(element, "filled")) &&
                                   (fillElement == null || IsNativeVmlSwitchEnabled(GetNativeOpenXmlAttribute(fillElement, "on")));
                if (fillEnabled) {
                    string? childFillColor = fillElement is not null ? GetNativeOpenXmlAttribute(fillElement, "color") : null;
                    string? fillColor = GetNativeOpenXmlAttribute(element, "fillcolor") ?? childFillColor;
                    bool explicitNoFill = IsNativeVmlNoColor(fillColor);
                    PdfCore.PdfColor? fill = explicitNoFill ? null : ParseNativeColor(NormalizeNativeVmlColor(fillColor));
                    if (!explicitNoFill) {
                        shape.FillColor = (fill ?? PdfCore.PdfColor.White).ToOfficeColor();
                    }

                    if (!explicitNoFill && TryGetNativeVmlGradientFill(fillElement, fill, out OfficeLinearGradient? gradient)) {
                        shape.FillGradient = gradient;
                    }

                    double? fillOpacity = fillElement is not null ? ParseNativeVmlOpacity(GetNativeOpenXmlAttribute(fillElement, "opacity")) : null;
                    if (fillOpacity.HasValue) {
                        shape.FillOpacity = fillOpacity.Value;
                    }
                }
            }

            ApplyNativeVmlShapeShadow(shape, element);
            OpenXmlElement? strokeElement = GetNativeVmlChild(element, "stroke");
            string? childStrokeColor = strokeElement is not null ? GetNativeOpenXmlAttribute(strokeElement, "color") : null;
            string? rawStrokeColor = GetNativeOpenXmlAttribute(element, "strokecolor") ?? childStrokeColor;
            string? strokeColor = NormalizeNativeVmlColor(rawStrokeColor);
            bool stroked = IsNativeVmlSwitchEnabled(GetNativeOpenXmlAttribute(element, "stroked")) &&
                           (strokeElement == null || IsNativeVmlSwitchEnabled(GetNativeOpenXmlAttribute(strokeElement, "on")));

            if (!stroked || IsNativeVmlNoColor(rawStrokeColor)) {
                shape.StrokeColor = null;
                shape.StrokeWidth = 0D;
                return;
            }

            shape.StrokeColor = (ParseNativeColor(strokeColor) ?? PdfCore.PdfColor.Black).ToOfficeColor();
            string? childStrokeWeight = strokeElement is not null ? GetNativeOpenXmlAttribute(strokeElement, "weight") : null;
            // VML defaults to a one-pixel stroke, which is 0.75 PDF points.
            shape.StrokeWidth = ParseNativeVmlStrokeWeight(GetNativeOpenXmlAttribute(element, "strokeweight") ?? childStrokeWeight) ?? 0.75D;
            shape.StrokeDashStyle = MapNativeVmlStrokeDashStyle(GetNativeOpenXmlAttribute(strokeElement ?? element, "dashstyle"));
            shape.StrokeLineCap = MapNativeVmlStrokeLineCap(GetNativeOpenXmlAttribute(strokeElement ?? element, "endcap"));
            shape.StrokeLineJoin = MapNativeVmlStrokeLineJoin(GetNativeOpenXmlAttribute(strokeElement ?? element, "joinstyle"));
            double? strokeOpacity = strokeElement is not null ? ParseNativeVmlOpacity(GetNativeOpenXmlAttribute(strokeElement, "opacity")) : null;
            if (strokeOpacity.HasValue) {
                shape.StrokeOpacity = strokeOpacity.Value;
            }
        }

        private static OfficeStrokeDashStyle MapNativeVmlStrokeDashStyle(string? value) {
            if (string.IsNullOrWhiteSpace(value)) {
                return OfficeStrokeDashStyle.Solid;
            }

            string normalized = value!.Replace(" ", string.Empty).Replace("-", string.Empty);
            bool hasDash = normalized.IndexOf("dash", StringComparison.OrdinalIgnoreCase) >= 0;
            bool hasDot = normalized.IndexOf("dot", StringComparison.OrdinalIgnoreCase) >= 0;
            if (hasDash && hasDot) {
                return OfficeStrokeDashStyle.DashDot;
            }

            if (hasDot) {
                return OfficeStrokeDashStyle.Dot;
            }

            return hasDash ? OfficeStrokeDashStyle.Dash : OfficeStrokeDashStyle.Solid;
        }

        private static OfficeStrokeLineCap? MapNativeVmlStrokeLineCap(string? value) {
            if (string.IsNullOrWhiteSpace(value)) {
                return null;
            }

            if (value!.Equals("round", StringComparison.OrdinalIgnoreCase)) {
                return OfficeStrokeLineCap.Round;
            }

            if (value.Equals("square", StringComparison.OrdinalIgnoreCase)) {
                return OfficeStrokeLineCap.Square;
            }

            return value.Equals("flat", StringComparison.OrdinalIgnoreCase) ||
                   value.Equals("butt", StringComparison.OrdinalIgnoreCase)
                ? OfficeStrokeLineCap.Butt
                : null;
        }

        private static OfficeStrokeLineJoin? MapNativeVmlStrokeLineJoin(string? value) {
            if (string.IsNullOrWhiteSpace(value)) {
                return null;
            }

            if (value!.Equals("round", StringComparison.OrdinalIgnoreCase)) {
                return OfficeStrokeLineJoin.Round;
            }

            if (value.Equals("bevel", StringComparison.OrdinalIgnoreCase)) {
                return OfficeStrokeLineJoin.Bevel;
            }

            return value.Equals("miter", StringComparison.OrdinalIgnoreCase)
                ? OfficeStrokeLineJoin.Miter
                : null;
        }

        private static void ApplyNativeVmlShapeShadow(OfficeShape shape, OpenXmlElement element) {
            OpenXmlElement? shadowElement = GetNativeVmlChild(element, "shadow");
            if (shadowElement == null ||
                !IsNativeVmlSwitchEnabled(GetNativeOpenXmlAttribute(shadowElement, "on"))) {
                return;
            }

            PdfCore.PdfColor shadowColor = ParseNativeColor(NormalizeNativeVmlColor(GetNativeOpenXmlAttribute(shadowElement, "color"))) ?? PdfCore.PdfColor.Black;
            double opacity = ParseNativeVmlOpacity(GetNativeOpenXmlAttribute(shadowElement, "opacity")) ?? 0.5D;
            (double offsetX, double offsetY) = ParseNativeVmlOffset(GetNativeOpenXmlAttribute(shadowElement, "offset")) ?? (2D, 2D);
            shape.Shadow = new OfficeShadow(shadowColor.ToOfficeColor(), opacity, offsetX, offsetY);
        }

        private static (double X, double Y)? ParseNativeVmlOffset(string? value) {
            if (string.IsNullOrWhiteSpace(value)) {
                return null;
            }

            string[] parts = value!.Split(',');
            if (parts.Length != 2) {
                return null;
            }

            double? x = ResolveNativeVmlLength(parts[0], 1D, 1D);
            double? y = ResolveNativeVmlLength(parts[1], 1D, 1D);
            return x.HasValue && y.HasValue ? (x.Value, y.Value) : null;
        }

        private static void ApplyNativeVmlShapeTransform(OfficeShape shape, OpenXmlElement element) {
            OfficeTransform? transform = null;

            string? flip = GetNativeVmlFlip(element);
            if (!string.IsNullOrWhiteSpace(flip)) {
                bool horizontal = flip!.IndexOf("x", StringComparison.OrdinalIgnoreCase) >= 0;
                bool vertical = flip.IndexOf("y", StringComparison.OrdinalIgnoreCase) >= 0;
                if (horizontal || vertical) {
                    OfficeTransform flipTransform = CreateNativeVmlCenterScaleTransform(
                        shape.Width,
                        shape.Height,
                        horizontal ? -1D : 1D,
                        vertical ? -1D : 1D);
                    transform = transform.HasValue ? transform.Value.Then(flipTransform) : flipTransform;
                }
            }

            double? rotation = GetNativeVmlRotationDegrees(element);
            if (rotation.HasValue && Math.Abs(rotation.Value) > 0.0001D) {
                OfficeTransform rotationTransform = OfficeTransform.RotateDegrees(rotation.Value, shape.Width / 2D, shape.Height / 2D);
                transform = transform.HasValue ? transform.Value.Then(rotationTransform) : rotationTransform;
            }

            if (transform.HasValue) {
                shape.Transform = shape.Transform.HasValue ? shape.Transform.Value.Then(transform.Value) : transform.Value;
            }
        }

        private static OfficeTransform CreateNativeVmlCenterScaleTransform(double width, double height, double scaleX, double scaleY) {
            double centerX = width / 2D;
            double centerY = height / 2D;
            return OfficeTransform.Translate(-centerX, -centerY)
                .Then(OfficeTransform.Scale(scaleX, scaleY))
                .Then(OfficeTransform.Translate(centerX, centerY));
        }

        private static double? GetNativeVmlRotationDegrees(OpenXmlElement element) {
            Dictionary<string, string> style = ParseNativeVmlStyle(GetNativeOpenXmlAttribute(element, "style"));
            string? value = style.TryGetValue("rotation", out string? styleRotation)
                ? styleRotation
                : GetNativeOpenXmlAttribute(element, "rotation");

            if (string.IsNullOrWhiteSpace(value)) {
                return null;
            }

            string normalized = value!.Trim();
            if (normalized.EndsWith("fd", StringComparison.OrdinalIgnoreCase)) {
                double? fixedPoint = ParseNativeVmlDouble(normalized.Substring(0, normalized.Length - 2));
                return fixedPoint.HasValue ? fixedPoint.Value / 65536D : null;
            }

            if (normalized.EndsWith("deg", StringComparison.OrdinalIgnoreCase)) {
                normalized = normalized.Substring(0, normalized.Length - 3);
            }

            return ParseNativeVmlDouble(normalized);
        }

        private static string? GetNativeVmlFlip(OpenXmlElement element) {
            Dictionary<string, string> style = ParseNativeVmlStyle(GetNativeOpenXmlAttribute(element, "style"));
            return style.TryGetValue("flip", out string? styleFlip)
                ? styleFlip
                : GetNativeOpenXmlAttribute(element, "flip");
        }

        private static bool TryGetNativeVmlGradientFill(OpenXmlElement? fillElement, PdfCore.PdfColor? startFill, out OfficeLinearGradient? gradient) {
            gradient = null;
            if (fillElement == null) {
                return false;
            }

            string? type = GetNativeOpenXmlAttribute(fillElement, "type");
            if (string.IsNullOrWhiteSpace(type) ||
                type!.IndexOf("gradient", StringComparison.OrdinalIgnoreCase) < 0) {
                return false;
            }

            string? color2Value = NormalizeNativeVmlColor(GetNativeOpenXmlAttribute(fillElement, "color2"));
            TryGetNativeVmlGradientStopColors(fillElement, out PdfCore.PdfColor? stopStartFill, out PdfCore.PdfColor? stopEndFill);
            PdfCore.PdfColor? gradientStartFill = startFill ?? stopStartFill;
            PdfCore.PdfColor? endFill = ParseNativeColor(color2Value);
            if (!endFill.HasValue) {
                endFill = stopEndFill;
            }

            if (!gradientStartFill.HasValue || !endFill.HasValue) {
                return false;
            }

            gradient = CreateNativeVmlLinearGradient(
                gradientStartFill.Value.ToOfficeColor(),
                endFill.Value.ToOfficeColor(),
                GetNativeOpenXmlAttribute(fillElement, "angle"));
            return true;
        }

        private static bool TryGetNativeVmlGradientStopColors(OpenXmlElement fillElement, out PdfCore.PdfColor? startFill, out PdfCore.PdfColor? endFill) {
            startFill = null;
            endFill = null;
            string? colors = GetNativeOpenXmlAttribute(fillElement, "colors");
            if (string.IsNullOrWhiteSpace(colors)) {
                return false;
            }

            var stops = new List<(double Offset, PdfCore.PdfColor Color)>();
            foreach (string entry in colors!.Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries)) {
                Match match = Regex.Match(entry.Trim(), "^(" + NativeVmlNumberPattern + "(?:%|f)?)\\s+(.+)$", RegexOptions.IgnoreCase);
                if (!match.Success) {
                    continue;
                }

                string offsetValue = match.Groups[1].Value;
                string offsetText = offsetValue.TrimEnd('%', 'f', 'F');
                if (!double.TryParse(offsetText, NumberStyles.Float, CultureInfo.InvariantCulture, out double offset)) {
                    continue;
                }

                if (offsetValue.EndsWith("%", StringComparison.OrdinalIgnoreCase)) {
                    offset /= 100D;
                } else if (offsetValue.EndsWith("f", StringComparison.OrdinalIgnoreCase)) {
                    offset /= 65536D;
                }

                string? colorValue = NormalizeNativeVmlColor(match.Groups[2].Value);
                PdfCore.PdfColor? color = ParseNativeColor(colorValue);
                if (!color.HasValue) {
                    continue;
                }

                stops.Add((Math.Max(0D, Math.Min(1D, offset)), color.Value));
            }

            if (stops.Count < 2) {
                return false;
            }

            stops.Sort((left, right) => left.Offset.CompareTo(right.Offset));
            startFill = stops[0].Color;
            endFill = stops[stops.Count - 1].Color;
            return true;
        }

        private static OfficeLinearGradient CreateNativeVmlLinearGradient(OfficeColor startColor, OfficeColor endColor, string? angleValue) {
            double angle = NormalizeNativeVmlAngleDegrees(ParseNativeVmlDouble(angleValue ?? string.Empty));
            if (IsNativeVmlAngleNear(angle, 0D) || IsNativeVmlAngleNear(angle, 180D)) {
                return OfficeLinearGradient.Horizontal(startColor, endColor);
            }

            return IsNativeVmlAngleNear(angle, 90D) || IsNativeVmlAngleNear(angle, 270D)
                ? OfficeLinearGradient.Vertical(startColor, endColor)
                : OfficeLinearGradient.DiagonalDown(startColor, endColor);
        }

        private static double NormalizeNativeVmlAngleDegrees(double? angle) {
            if (!angle.HasValue) {
                return 0D;
            }

            double degrees = angle.Value % 360D;
            return degrees < 0D ? degrees + 360D : degrees;
        }

        private static bool IsNativeVmlAngleNear(double angle, double target) {
            double distance = Math.Abs(angle - target);
            distance = Math.Min(distance, 360D - distance);
            return distance <= 22.5D;
        }

    }
}
