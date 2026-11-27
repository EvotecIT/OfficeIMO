using OfficeIMO.OpenXml.Internal;
using OfficeIMO.PowerPoint.LegacyPpt.Internal;
using OfficeIMO.PowerPoint.LegacyPpt.Model;
using OfficeIMO.PowerPoint.OpenXml;

namespace OfficeIMO.PowerPoint {
    public sealed partial class PowerPointDocument {
        private const int LegacyA = LegacyPptXml.A;

        /// <summary>Gives <c>p:cSld</c> the binary background, adding its picture to <paramref name="ownerPart"/>.</summary>
        private static void ApplyLegacyBackground(PowerPointPartRef ownerPart,
            OfficeXmlNode commonSlideData, LegacyPptBackground source) {
            OfficeXmlNode background = commonSlideData.CreateElement(LegacyPptXml.P, "bg");
            OfficeXmlNode properties = background.AddP("bgPr");
            if (!AppendLegacyBackgroundFill(ownerPart, properties, source)) {
                if (source.ForegroundColor != null) {
                    ApplyLegacyBackground(commonSlideData, source.ForegroundColor);
                }
                return;
            }
            ReplaceLegacyBackground(commonSlideData, background);
        }

        /// <summary>Gives <c>p:cSld</c> a solid background of one colour.</summary>
        private static void ApplyLegacyBackground(OfficeXmlNode commonSlideData, string color) {
            OfficeXmlNode background = commonSlideData.CreateElement(LegacyPptXml.P, "bg");
            background.AddP("bgPr").AddA("solidFill").AddA("srgbClr").Set("val", color);
            ReplaceLegacyBackground(commonSlideData, background);
        }

        private static void ReplaceLegacyBackground(OfficeXmlNode commonSlideData, OfficeXmlNode background) {
            OfficeXmlNode existing = commonSlideData.Child(LegacyPptXml.P, "bg");
            if (existing.IsNone) {
                commonSlideData.InsertOrdered(background, LegacyPptImportXml.CommonSlideDataOrder);
            } else {
                commonSlideData.InsertBefore(background, existing);
                existing.Remove();
            }
        }

        // Appends the background fill; false, adding nothing, when the binary background has no fill to project.
        private static bool AppendLegacyBackgroundFill(PowerPointPartRef ownerPart, OfficeXmlNode parent,
            LegacyPptBackground source) {
            switch (source.Kind) {
                case LegacyPptBackgroundKind.None:
                    parent.AddA("noFill");
                    return true;
                case LegacyPptBackgroundKind.Solid:
                    if (source.ForegroundColor == null) return false;
                    AppendLegacyBackgroundColor(parent.AddA("solidFill"), source.ForegroundColor,
                        source.ForegroundOpacity);
                    return true;
                case LegacyPptBackgroundKind.LinearGradient:
                case LegacyPptBackgroundKind.CenterGradient:
                case LegacyPptBackgroundKind.ShapeGradient:
                case LegacyPptBackgroundKind.ScaleGradient:
                case LegacyPptBackgroundKind.TitleGradient:
                    return AppendLegacyGradientFill(parent, source);
                case LegacyPptBackgroundKind.Pattern:
                case LegacyPptBackgroundKind.Texture:
                case LegacyPptBackgroundKind.Picture:
                    return AppendLegacyBackgroundBlipFill(ownerPart, parent, source);
                case LegacyPptBackgroundKind.Inherited:
                    parent.AddA("grpFill");
                    return true;
                default:
                    return false;
            }
        }

        private static bool AppendLegacyGradientFill(OfficeXmlNode parent,
            LegacyPptBackground source) {
            if (source.ForegroundColor == null || source.BackgroundColor == null) return false;
            OfficeXmlNode gradient = parent.AddA("gradFill").Set("rotWithShape", true);
            OfficeXmlNode stops = gradient.AddA("gsLst");
            LegacyPptGradientStop[] customStops = source.GradientStops
                .Where(stop => stop.Color != null)
                .ToArray();
            if (customStops.Length >= 2
                && customStops.Length == source.GradientStops.Count) {
                foreach (LegacyPptGradientStop stop in customStops) {
                    AppendLegacyGradientStop(stops, stop.Color!,
                        InterpolateLegacyGradientOpacity(source, stop.Position),
                        checked((int)Math.Round(stop.Position * 100000D, MidpointRounding.AwayFromZero)));
                }
            } else {
                AppendLegacyGradientStop(stops, source.ForegroundColor, source.ForegroundOpacity, 0);
                AppendLegacyGradientStop(stops, source.BackgroundColor, source.BackgroundOpacity, 100000);
            }
            if (source.Kind is LegacyPptBackgroundKind.LinearGradient
                    or LegacyPptBackgroundKind.ScaleGradient) {
                double angle = NormalizeLegacyGradientAngle(
                    270D - source.AngleDegrees.GetValueOrDefault());
                gradient.AddA("lin")
                    .Set("ang", checked((int)Math.Round(angle * 60000D, MidpointRounding.AwayFromZero)))
                    .Set("scaled", source.Kind == LegacyPptBackgroundKind.ScaleGradient);
            } else {
                gradient.AddA("path")
                    .Set("path", source.Kind == LegacyPptBackgroundKind.CenterGradient ? "circle" : "shape");
            }
            return true;
        }

        private static void AppendLegacyGradientStop(OfficeXmlNode stops, string color, double? opacity, int position) {
            OfficeXmlNode stop = stops.AddA("gs").Set("pos", position);
            AppendLegacyBackgroundColor(stop, color, opacity);
        }

        private static double? InterpolateLegacyGradientOpacity(
            LegacyPptBackground source, double position) {
            if (!source.ForegroundOpacity.HasValue && !source.BackgroundOpacity.HasValue) {
                return null;
            }
            double first = source.ForegroundOpacity ?? 1D;
            double last = source.BackgroundOpacity ?? first;
            return first + (last - first) * position;
        }

        /// <summary>
        /// A detached <c>a:gradFill</c> created under <paramref name="context"/>'s part for an OfficeArt gradient fill, or
        /// none when the shape has no gradient to project.
        /// </summary>
        private static OfficeXmlNode CreateLegacyShapeGradientFill(OfficeXmlNode context, LegacyPptShape source) {
            uint fillType = source.Style.FillType.GetValueOrDefault();
            if (fillType is < 4 or > 8) {
                return default;
            }

            LegacyPptGradientStop[] customStops = source.FillGradientStops
                .Where(stop => stop.Color != null)
                .ToArray();
            bool useCustomStops = customStops.Length >= 2
                && customStops.Length == source.FillGradientStops.Count;
            string? foreground = null;
            string? background = null;
            if (!useCustomStops) {
                foreground = source.FillColor
                    ?? (source.Style.FillColor.HasValue ? null : "FFFFFF");
                background = source.FillBackColor
                    ?? (source.Style.FillBackColor.HasValue ? null : "FFFFFF");
                if (foreground == null || background == null) {
                    return default;
                }
            }

            OfficeXmlNode gradient = context.CreateElement(LegacyA, "gradFill").Set("rotWithShape", false);
            OfficeXmlNode stops = gradient.AddA("gsLst");
            if (useCustomStops) {
                foreach (LegacyPptGradientStop stop in customStops) {
                    AppendLegacyGradientStop(stops, stop.Color!,
                        InterpolateLegacyShapeGradientOpacity(source, stop.Position),
                        checked((int)Math.Round(stop.Position * 100000D, MidpointRounding.AwayFromZero)));
                }
            } else {
                AppendLegacyGradientStop(stops, foreground!, source.Style.FillOpacity, 0);
                AppendLegacyGradientStop(stops, background!, source.Style.FillBackOpacity, 100000);
            }
            if (fillType is 4 or 7) {
                double angle = NormalizeLegacyGradientAngle(
                    90D - source.Style.FillAngleDegrees.GetValueOrDefault());
                gradient.AddA("lin")
                    .Set("ang", checked((int)Math.Round(angle * 60000D, MidpointRounding.AwayFromZero)))
                    .Set("scaled", fillType == 7);
            } else {
                gradient.AddA("path").Set("path", fillType == 5 ? "circle" : "shape");
            }
            return gradient;
        }

        private static double? InterpolateLegacyShapeGradientOpacity(
            LegacyPptShape source, double position) {
            if (!source.Style.FillOpacity.HasValue && !source.Style.FillBackOpacity.HasValue) {
                return null;
            }
            double first = source.Style.FillOpacity ?? 1D;
            double last = source.Style.FillBackOpacity ?? first;
            return first + (last - first) * position;
        }

        private static bool AppendLegacyBackgroundBlipFill(PowerPointPartRef ownerPart, OfficeXmlNode parent,
            LegacyPptBackground source) {
            if (source.Picture?.HasImportableImage != true) return false;
            string relationshipId = AddLegacyImagePart(ownerPart, source.Picture);
            OfficeXmlNode result = parent.AddA("blipFill");
            result.AddA("blip").SetRelationship("embed", relationshipId);
            if (source.Kind is LegacyPptBackgroundKind.Pattern
                    or LegacyPptBackgroundKind.Texture) {
                result.AddA("tile");
            } else {
                result.AddA("stretch").AddA("fillRect");
            }
            return true;
        }

        private static void AppendLegacyBackgroundColor(OfficeXmlNode parent, string color,
            double? opacity) {
            OfficeXmlNode result = parent.AddA("srgbClr").Set("val", color);
            if (opacity.HasValue) {
                result.AddA("alpha").Set("val",
                    checked((int)Math.Round(Math.Max(0D, Math.Min(1D, opacity.Value))
                        * 100000D, MidpointRounding.AwayFromZero)));
            }
        }

        private static double NormalizeLegacyGradientAngle(double angle) {
            double result = angle % 360D;
            return result < 0D ? result + 360D : result;
        }
    }
}
