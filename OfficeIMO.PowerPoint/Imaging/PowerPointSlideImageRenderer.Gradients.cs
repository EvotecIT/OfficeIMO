using OfficeIMO.Drawing;
using OfficeIMO.OpenXml.Internal;

namespace OfficeIMO.PowerPoint {
    internal static partial class PowerPointSlideImageRenderer {
        private enum ShapeFillGradientProjection {
            None,
            Applied,
            Unsupported
        }

        private static ShapeFillGradientProjection ApplyShapeFillGradient(
            OfficeShape target,
            PowerPointShape source,
            OfficeXmlNode colorScheme,
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
            OfficeXmlNode colorScheme,
            double width,
            double height,
            bool hasTransformedAncestor,
            out OfficeLinearGradient? linear,
            out OfficeRadialGradient? radial,
            out OfficeColor fallback) {
            linear = null;
            radial = null;
            fallback = default;
            if (!TryGetShapeFillGradient(source, out OfficeXmlNode gradient,
                    out OfficeXmlNode placeholderColor)) {
                return ShapeFillGradientProjection.None;
            }
            if (!TryResolveGradientStops(gradient, colorScheme, placeholderColor, GetSlideColorMap(source.OwnerSlide),
                    out IReadOnlyList<OfficeGradientStop>? stops)) {
                return ShapeFillGradientProjection.Unsupported;
            }

            fallback = stops![0].Color;
            bool? rotateWithShape = gradient.BooleanAttribute("rotWithShape");
            OfficeXmlNode linearFill = gradient.Child(D, "lin");
            if (!linearFill.IsNone) {
                double angle = (linearFill.Int32Attribute("ang") ?? 0) / 60000D;
                if (rotateWithShape == false) {
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

            OfficeXmlNode path = gradient.Child(D, "path");
            if (!path.IsNone
                && path.Attribute("path") == "circle"
                && path.Child(D, "fillToRect").IsNone
                && !hasTransformedAncestor
                && CanProjectCenteredRadialGradient(source, width, height,
                    rotateWithShape != false)) {
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
            out OfficeXmlNode gradient,
            out OfficeXmlNode placeholderColor) {
            placeholderColor = default;
            OfficeXmlNode properties = GetShapePropertiesElement(source);
            gradient = properties.Child(D, "gradFill");
            if (!gradient.IsNone) {
                return true;
            }
            if (!properties.IsNone && HasExplicitShapeFill(properties)) {
                return false;
            }

            OfficeXmlNode fillReference = GetShapeStyleElement(source).Child(D, "fillRef");
            uint? index = fillReference.UInt32Attribute("idx");
            OfficeXmlNode formatScheme = source.OwnerSlide == null
                ? default
                : GetSlideFormatScheme(source.OwnerSlide);
            OfficeXmlNode themeFill = ResolveThemeShapeFill(formatScheme, index);
            gradient = themeFill.Is(D, "gradFill") ? themeFill : themeFill.Child(D, "gradFill");
            if (gradient.IsNone) {
                return false;
            }

            placeholderColor = fillReference;
            return true;
        }

        /// <summary>The <c>p:style</c> of a shape or connector, or none for other elements.</summary>
        private static OfficeXmlNode GetShapeStyleElement(PowerPointShape source) {
            OfficeXmlNode element = source.XmlElement;
            return element.Is(Pml, "sp") || element.Is(Pml, "cxnSp") ? element.Child(Pml, "style") : default;
        }

        private static bool HasExplicitShapeFill(OfficeXmlNode properties) {
            foreach (OfficeXmlNode child in properties.Children()) {
                if (child.NamespaceId == D && child.LocalName is "noFill" or "solidFill" or "gradFill" or "blipFill" or "pattFill" or "grpFill") {
                    return true;
                }
            }
            return false;
        }

        private static OfficeXmlNode ResolveThemeShapeFill(OfficeXmlNode formatScheme,
            uint? index) {
            if (formatScheme.IsNone || !index.HasValue) {
                return default;
            }
            if (index.Value >= 1001U) {
                return ChildAt(formatScheme.Child(D, "bgFillStyleLst"), index.Value - 1001U);
            }
            if (index.Value < 1U) return default;
            return ChildAt(formatScheme.Child(D, "fillStyleLst"), index.Value - 1U);
        }

        private static OfficeXmlNode ResolveThemeShapeLine(
            OfficeXmlNode formatScheme, uint? index) {
            if (formatScheme.IsNone || !index.HasValue || index.Value < 1U) {
                return default;
            }
            return ChildAt(formatScheme.Child(D, "lnStyleLst"), index.Value - 1U);
        }

        /// <summary>The child element at a zero-based position, or none past the last child.</summary>
        private static OfficeXmlNode ChildAt(OfficeXmlNode parent, uint position) {
            uint current = 0U;
            foreach (OfficeXmlNode child in parent.Children()) {
                if (current++ == position) return child;
            }
            return default;
        }

        private static bool TryResolveGradientStops(OfficeXmlNode gradient,
            OfficeXmlNode colorScheme,
            OfficeXmlNode placeholderColor,
            OfficeXmlNode colorMap,
            out IReadOnlyList<OfficeGradientStop>? result) {
            result = null;
            OfficeXmlNode[] sourceStops = gradient.Child(D, "gsLst").Children(D, "gs").ToList()
                .OrderBy(stop => stop.Int32Attribute("pos") ?? 0)
                .ToArray();
            if (sourceStops.Length < 2) return false;

            var stops = new List<OfficeGradientStop>(sourceStops.Length + 2);
            foreach (OfficeXmlNode sourceStop in sourceStops) {
                OfficeColor? color = OfficeXmlThemeColors.ResolveColor(
                    sourceStop, colorScheme, placeholderColor, colorMap);
                if (!color.HasValue) return false;
                double offset = Math.Max(0D, Math.Min(1D,
                    (sourceStop.Int32Attribute("pos") ?? 0) / 100000D));
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
