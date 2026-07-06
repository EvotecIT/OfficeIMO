using System;
using System.Collections.Generic;
using System.Globalization;
using System.Xml.Linq;
using OfficeIMO.Drawing;

namespace OfficeIMO.Visio {
    internal static partial class VisioSvgPreviewRasterizer {
        private readonly struct SvgPaint {
            internal static SvgPaint Default => new(OfficeColor.Black, null, null, OfficeColor.Transparent, 1D, null, 1D, OfficeColor.Black, SvgStrokeLineCap.Butt, SvgStrokeLineJoin.Miter, false);

            private SvgPaint(OfficeColor fill, OfficeLinearGradient? fillGradient, OfficeRadialGradient? fillRadialGradient, OfficeColor stroke, double strokeWidth, IReadOnlyList<double>? dashPattern, double opacity, OfficeColor currentColor, SvgStrokeLineCap strokeLineCap, SvgStrokeLineJoin strokeLineJoin, bool nonScalingStroke) {
                Fill = fill;
                FillGradient = fillGradient;
                FillRadialGradient = fillRadialGradient;
                Stroke = stroke;
                StrokeWidth = strokeWidth;
                DashPattern = dashPattern;
                Opacity = opacity;
                CurrentColor = currentColor;
                StrokeLineCap = strokeLineCap;
                StrokeLineJoin = strokeLineJoin;
                NonScalingStroke = nonScalingStroke;
            }

            internal OfficeColor Fill { get; }

            internal OfficeLinearGradient? FillGradient { get; }

            internal OfficeRadialGradient? FillRadialGradient { get; }

            internal OfficeColor Stroke { get; }

            internal double StrokeWidth { get; }

            internal IReadOnlyList<double>? DashPattern { get; }

            internal double Opacity { get; }

            internal OfficeColor CurrentColor { get; }

            internal SvgStrokeLineCap StrokeLineCap { get; }

            internal SvgStrokeLineJoin StrokeLineJoin { get; }

            internal bool NonScalingStroke { get; }

            internal bool HasFill => Fill.A > 0 || FillGradient != null || FillRadialGradient != null;

            internal static SvgPaint Resolve(XElement element, SvgPaint inherited, SvgRenderContext context) {
                Dictionary<string, string> style = context.StyleSheet.CreateStyle(element);
                OfficeColor currentColor = ResolveColor(ReadPaint(element, style, "color"), inherited.CurrentColor, inherited.CurrentColor);
                string? rawFill = ReadPaint(element, style, "fill");
                string? rawStroke = ReadPaint(element, style, "stroke");
                double ownOpacity = ReadOwnUnit(element, style, "opacity", 1D);
                double fillOpacity = ReadUnit(element, style, "fill-opacity", 1D);
                double strokeOpacity = ReadUnit(element, style, "stroke-opacity", 1D);
                double explicitFillMultiplier = string.IsNullOrWhiteSpace(rawFill) ? 1D : inherited.Opacity;
                double explicitStrokeMultiplier = string.IsNullOrWhiteSpace(rawStroke) ? 1D : inherited.Opacity;
                OfficeColor fill = ResolveColor(rawFill, inherited.Fill, currentColor);
                OfficeColor stroke = ResolveColor(rawStroke, inherited.Stroke, currentColor);
                double opacity = inherited.Opacity * ownOpacity;
                fill = ApplyAlpha(fill, explicitFillMultiplier * ownOpacity * fillOpacity);
                stroke = ApplyAlpha(stroke, explicitStrokeMultiplier * ownOpacity * strokeOpacity);
                OfficeLinearGradient? fillGradient = string.IsNullOrWhiteSpace(rawFill) ? inherited.FillGradient : null;
                OfficeRadialGradient? fillRadialGradient = string.IsNullOrWhiteSpace(rawFill) ? inherited.FillRadialGradient : null;
                if (TryResolveFillGradient(rawFill, context, explicitFillMultiplier * ownOpacity * fillOpacity, currentColor, out OfficeLinearGradient? resolvedFillGradient, out OfficeRadialGradient? resolvedFillRadialGradient)) {
                    fillGradient = resolvedFillGradient;
                    fillRadialGradient = resolvedFillRadialGradient;
                }

                double strokeWidth = ReadLength(element, "stroke-width", inherited.StrokeWidth);
                if (style.TryGetValue("stroke-width", out string? styleStrokeWidth) && TryParseLength(styleStrokeWidth, out double parsedStrokeWidth)) {
                    strokeWidth = parsedStrokeWidth;
                }

                IReadOnlyList<double>? dashPattern = ReadDashPattern(element, style, inherited.DashPattern);
                SvgStrokeLineCap strokeLineCap = ReadStrokeLineCap(element, style, inherited.StrokeLineCap);
                SvgStrokeLineJoin strokeLineJoin = ReadStrokeLineJoin(element, style, inherited.StrokeLineJoin);
                bool nonScalingStroke = ReadNonScalingStroke(element, style, inherited.NonScalingStroke);
                return new SvgPaint(fill, fillGradient, fillRadialGradient, stroke, strokeWidth, dashPattern, opacity, currentColor, strokeLineCap, strokeLineJoin, nonScalingStroke);
            }

            private static string? ReadPaint(XElement element, Dictionary<string, string> style, string name) =>
                ReadAttributeOrStyle(element, style, name);

            private static double ReadUnit(XElement element, Dictionary<string, string> style, string name, double fallback) {
                string? raw = ReadAttributeOrStyle(element, style, name);
                if (string.IsNullOrWhiteSpace(raw) || !double.TryParse(raw, NumberStyles.Float, CultureInfo.InvariantCulture, out double parsed)) {
                    return fallback;
                }

                return Math.Max(0D, Math.Min(1D, parsed));
            }

            private static double ReadOwnUnit(XElement element, Dictionary<string, string> style, string name, double fallback) =>
                ReadUnit(element, style, name, fallback);

            private static IReadOnlyList<double>? ReadDashPattern(XElement element, Dictionary<string, string> style, IReadOnlyList<double>? inherited) {
                string? raw = ReadAttributeOrStyle(element, style, "stroke-dasharray");
                if (string.IsNullOrWhiteSpace(raw)) {
                    return inherited;
                }

                string trimmed = raw!.Trim();
                if (string.Equals(trimmed, "none", StringComparison.OrdinalIgnoreCase)) {
                    return null;
                }

                if (!TryParseNumbers(trimmed, out List<double> values) || values.Count == 0) {
                    return inherited;
                }

                List<double> pattern = new(values.Count * 2);
                for (int i = 0; i < values.Count; i++) {
                    if (values[i] > 0D && !double.IsNaN(values[i]) && !double.IsInfinity(values[i])) {
                        pattern.Add(values[i]);
                    }
                }

                if (pattern.Count == 0) {
                    return null;
                }

                if ((pattern.Count & 1) == 1) {
                    int count = pattern.Count;
                    for (int i = 0; i < count; i++) {
                        pattern.Add(pattern[i]);
                    }
                }

                return pattern;
            }

            private static SvgStrokeLineCap ReadStrokeLineCap(XElement element, Dictionary<string, string> style, SvgStrokeLineCap inherited) {
                string? raw = ReadAttributeOrStyle(element, style, "stroke-linecap");
                if (string.IsNullOrWhiteSpace(raw)) {
                    return inherited;
                }

                return raw!.Trim().ToLowerInvariant() switch {
                    "round" => SvgStrokeLineCap.Round,
                    "square" => SvgStrokeLineCap.Square,
                    _ => SvgStrokeLineCap.Butt
                };
            }

            private static SvgStrokeLineJoin ReadStrokeLineJoin(XElement element, Dictionary<string, string> style, SvgStrokeLineJoin inherited) {
                string? raw = ReadAttributeOrStyle(element, style, "stroke-linejoin");
                if (string.IsNullOrWhiteSpace(raw)) {
                    return inherited;
                }

                return raw!.Trim().ToLowerInvariant() switch {
                    "round" => SvgStrokeLineJoin.Round,
                    "bevel" => SvgStrokeLineJoin.Bevel,
                    _ => SvgStrokeLineJoin.Miter
                };
            }

            private static bool ReadNonScalingStroke(XElement element, Dictionary<string, string> style, bool inherited) {
                string? raw = ReadAttributeOrStyle(element, style, "vector-effect");
                if (string.IsNullOrWhiteSpace(raw)) {
                    return inherited;
                }

                string vectorEffect = raw!.Trim();
                if (string.Equals(vectorEffect, "none", StringComparison.OrdinalIgnoreCase)) {
                    return false;
                }

                return vectorEffect.IndexOf("non-scaling-stroke", StringComparison.OrdinalIgnoreCase) >= 0;
            }

            private static string? ReadAttributeOrStyle(XElement element, Dictionary<string, string> style, string name) =>
                style.TryGetValue(name, out string? value) ? value : element.Attribute(name)?.Value;

            private static OfficeColor ResolveColor(string? raw, OfficeColor inherited, OfficeColor currentColor) {
                if (string.IsNullOrWhiteSpace(raw)) {
                    return inherited;
                }

                string value = raw!.Trim();
                if (string.Equals(value, "none", StringComparison.OrdinalIgnoreCase)) {
                    return OfficeColor.Transparent;
                }

                if (string.Equals(value, "currentColor", StringComparison.OrdinalIgnoreCase)) {
                    return currentColor;
                }

                if (value.StartsWith("rgb(", StringComparison.OrdinalIgnoreCase) && value.EndsWith(")", StringComparison.Ordinal)) {
                    string inner = value.Substring(4, value.Length - 5);
                    if (TryParseNumbers(inner, out List<double> parts) && parts.Count >= 3) {
                        return OfficeColor.FromRgb(ToByte(parts[0]), ToByte(parts[1]), ToByte(parts[2]));
                    }
                }

                if (OfficeColor.TryParse(value, out OfficeColor color)) {
                    return color;
                }

                return inherited;
            }

            private static OfficeColor ApplyAlpha(OfficeColor color, double opacity) =>
                OfficeColor.FromRgba(color.R, color.G, color.B, (byte)Math.Round(color.A * Math.Max(0D, Math.Min(1D, opacity))));

            private static byte ToByte(double value) => (byte)Math.Max(0D, Math.Min(255D, Math.Round(value)));
        }

        private enum SvgStrokeLineCap {
            Butt,
            Round,
            Square
        }

        private enum SvgStrokeLineJoin {
            Miter,
            Round,
            Bevel
        }
    }
}
