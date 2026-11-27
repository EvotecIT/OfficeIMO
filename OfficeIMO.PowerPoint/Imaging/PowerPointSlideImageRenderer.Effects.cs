using OfficeIMO.Drawing;
using OfficeIMO.OpenXml.Internal;

namespace OfficeIMO.PowerPoint {
    internal static partial class PowerPointSlideImageRenderer {
        private static void ApplyShapeEffects(OfficeShape target, PowerPointShape source, OfficeXmlNode colorScheme, PowerPointShapeBoundsMapping mapping) {
            OfficeXmlNode effects = GetShapePropertiesElement(source).Child(D, "effectLst");
            if (effects.IsNone) {
                return;
            }

            OfficeXmlNode glow = effects.Child(D, "glow");
            if (!glow.IsNone && TryCreateGlow(glow, colorScheme, mapping, out OfficeGlow? officeGlow, GetSlideColorMap(source.OwnerSlide))) {
                target.Glow = officeGlow;
            }

            OfficeXmlNode shadow = effects.Child(D, "outerShdw");
            if (!shadow.IsNone && TryCreateShadow(shadow, colorScheme, mapping, out OfficeShadow? officeShadow, GetSlideColorMap(source.OwnerSlide))) {
                target.Shadow = officeShadow;
            }
        }

        private static bool TryCreateGlow(OfficeXmlNode glow, OfficeXmlNode colorScheme, PowerPointShapeBoundsMapping mapping, out OfficeGlow? officeGlow, OfficeXmlNode colorMap = default) {
            officeGlow = null;
            double radius = mapping.MapHorizontalLength(PowerPointUnits.ToPoints(glow.Int64Attribute("rad") ?? 0L));
            if (radius <= 0D) {
                return false;
            }

            if (!TryResolveEffectColor(glow, colorScheme, colorMap, out OfficeColor color)) {
                return false;
            }

            double opacity = color.A / 255D;
            if (opacity <= 0D) {
                return false;
            }

            officeGlow = new OfficeGlow(OfficeColor.FromRgb(color.R, color.G, color.B), opacity, radius);
            return true;
        }

        private static bool TryCreateShadow(OfficeXmlNode shadow, OfficeXmlNode colorScheme, PowerPointShapeBoundsMapping mapping, out OfficeShadow? officeShadow, OfficeXmlNode colorMap = default) {
            officeShadow = null;
            double distance = mapping.MapHorizontalLength(PowerPointUnits.ToPoints(shadow.Int64Attribute("dist") ?? 0L));
            double blurRadius = mapping.MapHorizontalLength(PowerPointUnits.ToPoints(shadow.Int64Attribute("blurRad") ?? 0L));
            if (distance <= 0D && blurRadius <= 0D) {
                return false;
            }

            if (!TryResolveEffectColor(shadow, colorScheme, colorMap, out OfficeColor color)) {
                return false;
            }

            double opacity = color.A / 255D;
            if (opacity <= 0D) {
                return false;
            }

            double angleDegrees = (shadow.Int32Attribute("dir") ?? 0) / DrawingMlAngleUnitsPerDegree;
            double radians = angleDegrees * Math.PI / 180D;
            double offsetX = Math.Cos(radians) * distance;
            double offsetY = Math.Sin(radians) * distance;
            OfficeColor shadowColor = OfficeColor.FromRgb(color.R, color.G, color.B);
            officeShadow = new OfficeShadow(shadowColor, opacity, offsetX, offsetY, blurRadius);
            return true;
        }

        private static bool TryResolveEffectColor(OfficeXmlNode owner, OfficeXmlNode colorScheme,
            OfficeXmlNode colorMap, out OfficeColor color) {
            OfficeColor? resolved = OfficeXmlThemeColors.ResolveColor(owner, colorScheme, colorMap: colorMap);
            color = resolved.GetValueOrDefault();
            return resolved.HasValue;
        }
    }
}
