using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using OfficeIMO.Drawing;
using OfficeIMO.OpenXml.Internal;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace OfficeIMO.PowerPoint {
    internal static partial class PowerPointSlideImageRenderer {
        private enum ShapeFillGradientProjection {
            None,
            Applied,
            Unsupported
        }

        private static bool HasShapeFillGradient(PowerPointShape source) =>
            TryGetShapeFillGradient(source, out _, out _);

        private static ShapeFillGradientProjection ApplyShapeFillGradient(
            OfficeShape target,
            PowerPointShape source,
            A.ColorScheme? colorScheme,
            bool hasTransformedAncestor) {
            ShapeFillGradientProjection projection = TryCreateShapeFillGradient(
                source,
                colorScheme,
                target.Width,
                target.Height,
                hasTransformedAncestor,
                out OfficeLinearGradient? linear,
                out OfficeRadialGradient? radial,
                out OfficeColor fallback);
            if (projection != ShapeFillGradientProjection.Applied) {
                return projection;
            }

            target.FillColor = fallback;
            target.FillGradient = linear;
            target.FillRadialGradient = radial;
            return ShapeFillGradientProjection.Applied;
        }

        private static ShapeFillGradientProjection TryCreateShapeFillGradient(
            PowerPointShape source,
            A.ColorScheme? colorScheme,
            double width,
            double height,
            bool hasTransformedAncestor,
            out OfficeLinearGradient? linear,
            out OfficeRadialGradient? radial,
            out OfficeColor fallback) {
            linear = null;
            radial = null;
            fallback = default;
            if (!TryGetShapeFillGradient(source, out A.GradientFill? gradient,
                    out A.SchemeColor? placeholderColor)) {
                return ShapeFillGradientProjection.None;
            }
            if (!TryResolveGradientStops(gradient!, colorScheme, placeholderColor,
                    out IReadOnlyList<OfficeGradientStop>? stops)) {
                return ShapeFillGradientProjection.Unsupported;
            }

            fallback = stops![0].Color;
            A.LinearGradientFill? linearFill = gradient!.GetFirstChild<A.LinearGradientFill>();
            if (linearFill != null) {
                double angle = (linearFill.Angle?.Value ?? 0) / 60000D;
                if (gradient.RotateWithShape?.Value == false) {
                    if (hasTransformedAncestor) {
                        return ShapeFillGradientProjection.Unsupported;
                    }
                    var frameTransform = new OfficeImageFrameTransform(
                        source.Rotation ?? 0D,
                        width / 2D,
                        height / 2D,
                        source.HorizontalFlip == true,
                        source.VerticalFlip == true);
                    linear = OfficeLinearGradient.FromTransformedAngle(stops,
                        angle, width, height,
                        frameTransform.CreateDestinationTransform());
                } else {
                    linear = OfficeLinearGradient.FromAngle(stops, angle);
                }
                return ShapeFillGradientProjection.Applied;
            }

            A.PathGradientFill? path = gradient.GetFirstChild<A.PathGradientFill>();
            if (path != null
                && path.Path?.Value == A.PathShadeValues.Circle
                && path.GetFirstChild<A.FillToRectangle>() == null
                && !hasTransformedAncestor
                && CanProjectCenteredRadialGradient(source, width, height,
                    gradient.RotateWithShape?.Value != false)) {
                radial = new OfficeRadialGradient(
                    0.5D, 0.5D, 0D,
                    0.5D, 0.5D, 0.5D,
                    stops);
                return ShapeFillGradientProjection.Applied;
            }

            return ShapeFillGradientProjection.Unsupported;
        }

        private static bool CanProjectCenteredRadialGradient(PowerPointShape source,
            double width, double height, bool rotateWithShape) {
            if (Math.Abs(width - height) < 0.000001D) {
                return true;
            }
            double rotation = NormalizeGradientRotation(source.Rotation ?? 0D);
            double interval = rotateWithShape ? 90D : 180D;
            return Math.Abs(rotation / interval
                - Math.Round(rotation / interval)) < 0.000001D;
        }

        private static double NormalizeGradientRotation(double degrees) {
            double normalized = degrees % 360D;
            return normalized < 0D ? normalized + 360D : normalized;
        }

        private static bool TryGetShapeFillGradient(PowerPointShape source,
            out A.GradientFill? gradient,
            out A.SchemeColor? placeholderColor) {
            gradient = null;
            placeholderColor = null;
            P.ShapeProperties? properties = GetOpenXmlShapeProperties(source);
            gradient = properties?.GetFirstChild<A.GradientFill>();
            if (gradient != null) {
                return true;
            }
            if (properties != null && HasExplicitShapeFill(properties)) {
                return false;
            }

            A.FillReference? fillReference = GetOpenXmlShapeStyle(source)?.FillReference;
            uint? index = fillReference?.Index?.Value;
            A.FormatScheme? formatScheme = source.OwnerSlide == null
                ? null
                : GetSlideFormatScheme(source.OwnerSlide);
            OpenXmlElement? themeFill = ResolveThemeShapeFill(formatScheme, index);
            gradient = themeFill as A.GradientFill
                ?? themeFill?.GetFirstChild<A.GradientFill>();
            if (gradient == null) {
                return false;
            }

            placeholderColor = fillReference?.GetFirstChild<A.SchemeColor>();
            return true;
        }

        private static P.ShapeStyle? GetOpenXmlShapeStyle(PowerPointShape source) =>
            source.Element switch {
                P.Shape shape => shape.ShapeStyle,
                P.ConnectionShape connector => connector.ShapeStyle,
                _ => null
            };

        private static bool HasExplicitShapeFill(P.ShapeProperties properties) =>
            properties.ChildElements.Any(child => child is A.NoFill
                or A.SolidFill
                or A.GradientFill
                or A.BlipFill
                or A.PatternFill
                or A.GroupFill);

        private static OpenXmlElement? ResolveThemeShapeFill(A.FormatScheme? formatScheme,
            uint? index) {
            if (formatScheme == null || !index.HasValue) {
                return null;
            }
            if (index.Value >= 1001U) {
                OpenXmlElementList fills = formatScheme
                    .GetFirstChild<A.BackgroundFillStyleList>()?
                    .ChildElements ?? default;
                uint zeroBased = index.Value - 1001U;
                return zeroBased < unchecked((uint)fills.Count)
                    ? fills[unchecked((int)zeroBased)]
                    : null;
            }
            if (index.Value < 1U) return null;
            OpenXmlElementList styles = formatScheme
                .GetFirstChild<A.FillStyleList>()?.ChildElements ?? default;
            uint styleIndex = index.Value - 1U;
            return styleIndex < unchecked((uint)styles.Count)
                ? styles[unchecked((int)styleIndex)]
                : null;
        }

        private static bool TryResolveGradientStops(A.GradientFill gradient,
            A.ColorScheme? colorScheme,
            A.SchemeColor? placeholderColor,
            out IReadOnlyList<OfficeGradientStop>? result) {
            result = null;
            A.GradientStop[] sourceStops = gradient.GetFirstChild<A.GradientStopList>()?
                .Elements<A.GradientStop>()
                .OrderBy(stop => stop.Position?.Value ?? 0)
                .ToArray() ?? Array.Empty<A.GradientStop>();
            if (sourceStops.Length < 2) return false;

            var stops = new List<OfficeGradientStop>(sourceStops.Length + 2);
            foreach (A.GradientStop sourceStop in sourceStops) {
                OfficeColor? color = OfficeOpenXmlThemeColorResolver.ResolveColor(
                    sourceStop, colorScheme, placeholderColor);
                if (!color.HasValue) return false;
                double offset = Math.Max(0D, Math.Min(1D,
                    (sourceStop.Position?.Value ?? 0) / 100000D));
                stops.Add(new OfficeGradientStop(offset, color.Value));
            }

            if (stops[0].Offset > 0D) {
                stops.Insert(0, new OfficeGradientStop(0D, stops[0].Color));
            }
            if (stops[stops.Count - 1].Offset < 1D) {
                stops.Add(new OfficeGradientStop(1D, stops[stops.Count - 1].Color));
            }
            result = stops;
            return true;
        }
    }
}
