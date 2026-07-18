using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using OfficeIMO.PowerPoint.LegacyPpt.Model;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.PowerPoint {
    public sealed partial class PowerPointPresentation {
        private static void ApplyLegacyBackground(OpenXmlPart ownerPart,
            CommonSlideData commonSlideData, LegacyPptBackground source) {
            OpenXmlElement? fill = CreateLegacyBackgroundFill(ownerPart, source);
            if (fill == null) {
                if (source.ForegroundColor != null) {
                    ApplyLegacyBackground(commonSlideData, source.ForegroundColor);
                }
                return;
            }
            commonSlideData.Background = new Background(new BackgroundProperties(fill));
        }

        private static OpenXmlElement? CreateLegacyBackgroundFill(OpenXmlPart ownerPart,
            LegacyPptBackground source) {
            switch (source.Kind) {
                case LegacyPptBackgroundKind.None:
                    return new A.NoFill();
                case LegacyPptBackgroundKind.Solid:
                    return source.ForegroundColor == null
                        ? null
                        : new A.SolidFill(CreateLegacyBackgroundColor(source.ForegroundColor,
                            source.ForegroundOpacity));
                case LegacyPptBackgroundKind.LinearGradient:
                case LegacyPptBackgroundKind.CenterGradient:
                case LegacyPptBackgroundKind.ShapeGradient:
                case LegacyPptBackgroundKind.ScaleGradient:
                case LegacyPptBackgroundKind.TitleGradient:
                    return CreateLegacyGradientFill(source);
                case LegacyPptBackgroundKind.Pattern:
                case LegacyPptBackgroundKind.Texture:
                case LegacyPptBackgroundKind.Picture:
                    return CreateLegacyBackgroundBlipFill(ownerPart, source);
                case LegacyPptBackgroundKind.Inherited:
                    return new A.GroupFill();
                default:
                    return null;
            }
        }

        private static A.GradientFill? CreateLegacyGradientFill(
            LegacyPptBackground source) {
            if (source.ForegroundColor == null || source.BackgroundColor == null) return null;
            var gradient = new A.GradientFill { RotateWithShape = true };
            var stops = new A.GradientStopList();
            LegacyPptGradientStop[] customStops = source.GradientStops
                .Where(stop => stop.Color != null)
                .ToArray();
            if (customStops.Length >= 2
                && customStops.Length == source.GradientStops.Count) {
                foreach (LegacyPptGradientStop stop in customStops) {
                    stops.Append(new A.GradientStop(CreateLegacyBackgroundColor(stop.Color!,
                        InterpolateLegacyGradientOpacity(source, stop.Position))) {
                        Position = checked((int)Math.Round(stop.Position * 100000D,
                            MidpointRounding.AwayFromZero))
                    });
                }
            } else {
                stops.Append(
                    new A.GradientStop(CreateLegacyBackgroundColor(source.ForegroundColor,
                        source.ForegroundOpacity)) { Position = 0 },
                    new A.GradientStop(CreateLegacyBackgroundColor(source.BackgroundColor,
                        source.BackgroundOpacity)) { Position = 100000 });
            }
            gradient.Append(stops);
            if (source.Kind is LegacyPptBackgroundKind.LinearGradient
                    or LegacyPptBackgroundKind.ScaleGradient) {
                double angle = NormalizeLegacyGradientAngle(
                    270D - source.AngleDegrees.GetValueOrDefault());
                gradient.Append(new A.LinearGradientFill {
                    Angle = checked((int)Math.Round(angle * 60000D,
                        MidpointRounding.AwayFromZero)),
                    Scaled = source.Kind == LegacyPptBackgroundKind.ScaleGradient
                });
            } else {
                gradient.Append(new A.PathGradientFill {
                    Path = source.Kind == LegacyPptBackgroundKind.CenterGradient
                        ? A.PathShadeValues.Circle
                        : A.PathShadeValues.Shape
                });
            }
            return gradient;
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

        private static A.GradientFill? CreateLegacyShapeGradientFill(LegacyPptShape source) {
            uint fillType = source.Style.FillType.GetValueOrDefault();
            if (fillType is < 4 or > 8) {
                return null;
            }

            var gradient = new A.GradientFill { RotateWithShape = false };
            var stops = new A.GradientStopList();
            LegacyPptGradientStop[] customStops = source.FillGradientStops
                .Where(stop => stop.Color != null)
                .ToArray();
            if (customStops.Length >= 2
                && customStops.Length == source.FillGradientStops.Count) {
                foreach (LegacyPptGradientStop stop in customStops) {
                    stops.Append(new A.GradientStop(CreateLegacyBackgroundColor(stop.Color!,
                        InterpolateLegacyShapeGradientOpacity(source, stop.Position))) {
                        Position = checked((int)Math.Round(stop.Position * 100000D,
                            MidpointRounding.AwayFromZero))
                    });
                }
            } else {
                string? foreground = source.FillColor
                    ?? (source.Style.FillColor.HasValue ? null : "FFFFFF");
                string? background = source.FillBackColor
                    ?? (source.Style.FillBackColor.HasValue ? null : "FFFFFF");
                if (foreground == null || background == null) {
                    return null;
                }
                stops.Append(
                    new A.GradientStop(CreateLegacyBackgroundColor(foreground,
                        source.Style.FillOpacity)) { Position = 0 },
                    new A.GradientStop(CreateLegacyBackgroundColor(background,
                        source.Style.FillBackOpacity)) { Position = 100000 });
            }
            gradient.Append(stops);
            if (fillType is 4 or 7) {
                double angle = NormalizeLegacyGradientAngle(
                    90D - source.Style.FillAngleDegrees.GetValueOrDefault());
                gradient.Append(new A.LinearGradientFill {
                    Angle = checked((int)Math.Round(angle * 60000D,
                        MidpointRounding.AwayFromZero)),
                    Scaled = fillType == 7
                });
            } else {
                gradient.Append(new A.PathGradientFill {
                    Path = fillType == 5
                        ? A.PathShadeValues.Circle
                        : A.PathShadeValues.Shape
                });
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

        private static A.BlipFill? CreateLegacyBackgroundBlipFill(OpenXmlPart ownerPart,
            LegacyPptBackground source) {
            if (source.Picture?.HasImportableImage != true) return null;
            ImagePart imagePart = AddLegacyImagePart(ownerPart, source.Picture);
            string relationshipId = ownerPart.GetIdOfPart(imagePart);
            var result = new A.BlipFill(new A.Blip { Embed = relationshipId });
            if (source.Kind is LegacyPptBackgroundKind.Pattern
                    or LegacyPptBackgroundKind.Texture) {
                result.Append(new A.Tile());
            } else {
                result.Append(new A.Stretch(new A.FillRectangle()));
            }
            return result;
        }

        private static A.RgbColorModelHex CreateLegacyBackgroundColor(string color,
            double? opacity) {
            var result = new A.RgbColorModelHex { Val = color };
            if (opacity.HasValue) {
                result.Append(new A.Alpha {
                    Val = checked((int)Math.Round(Math.Max(0D, Math.Min(1D, opacity.Value))
                        * 100000D, MidpointRounding.AwayFromZero))
                });
            }
            return result;
        }

        private static double NormalizeLegacyGradientAngle(double angle) {
            double result = angle % 360D;
            return result < 0D ? result + 360D : result;
        }
    }
}
