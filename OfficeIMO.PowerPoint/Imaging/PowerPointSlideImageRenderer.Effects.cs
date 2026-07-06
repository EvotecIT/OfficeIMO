using System;
using DocumentFormat.OpenXml;
using OfficeIMO.Drawing;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.PowerPoint {
    internal static partial class PowerPointSlideImageRenderer {
        private static void ApplyShapeEffects(OfficeShape target, PowerPointShape source, A.ColorScheme? colorScheme, PowerPointShapeBoundsMapping mapping) {
            A.EffectList? effects = GetOpenXmlShapeProperties(source)?.GetFirstChild<A.EffectList>();
            if (effects == null) {
                return;
            }

            A.Glow? glow = effects.GetFirstChild<A.Glow>();
            if (glow != null && TryCreateGlow(glow, colorScheme, mapping, out OfficeGlow? officeGlow)) {
                target.Glow = officeGlow;
            }

            A.OuterShadow? shadow = effects.GetFirstChild<A.OuterShadow>();
            if (shadow != null && TryCreateShadow(shadow, colorScheme, mapping, out OfficeShadow? officeShadow)) {
                target.Shadow = officeShadow;
            }
        }

        private static bool TryCreateGlow(A.Glow glow, A.ColorScheme? colorScheme, PowerPointShapeBoundsMapping mapping, out OfficeGlow? officeGlow) {
            officeGlow = null;
            double radius = mapping.MapHorizontalLength(PowerPointUnits.ToPoints(glow.Radius?.Value ?? 0L));
            if (radius <= 0D) {
                return false;
            }

            if (!TryResolveEffectColor(glow, colorScheme, out OfficeColor color)) {
                return false;
            }

            double opacity = color.A / 255D;
            if (opacity <= 0D) {
                return false;
            }

            officeGlow = new OfficeGlow(OfficeColor.FromRgb(color.R, color.G, color.B), opacity, radius);
            return true;
        }

        private static bool TryCreateShadow(A.OuterShadow shadow, A.ColorScheme? colorScheme, PowerPointShapeBoundsMapping mapping, out OfficeShadow? officeShadow) {
            officeShadow = null;
            double distance = mapping.MapHorizontalLength(PowerPointUnits.ToPoints(shadow.Distance?.Value ?? 0L));
            if (distance <= 0D) {
                return false;
            }

            if (!TryResolveEffectColor(shadow, colorScheme, out OfficeColor color)) {
                return false;
            }

            double opacity = color.A / 255D;
            if (opacity <= 0D) {
                return false;
            }

            double blurRadius = mapping.MapHorizontalLength(PowerPointUnits.ToPoints(shadow.BlurRadius?.Value ?? 0L));
            double angleDegrees = (shadow.Direction?.Value ?? 0) / DrawingMlAngleUnitsPerDegree;
            double radians = angleDegrees * Math.PI / 180D;
            double offsetX = Math.Cos(radians) * distance;
            double offsetY = Math.Sin(radians) * distance;
            OfficeColor shadowColor = OfficeColor.FromRgb(color.R, color.G, color.B);
            officeShadow = new OfficeShadow(shadowColor, opacity, offsetX, offsetY, blurRadius);
            return true;
        }

        private static bool TryResolveEffectColor(OpenXmlCompositeElement owner, A.ColorScheme? colorScheme, out OfficeColor color) {
            color = default;
            A.SolidFill solidFill = new A.SolidFill();
            OpenXmlElement? colorElement = owner.GetFirstChild<A.RgbColorModelHex>();
            if (colorElement == null) {
                colorElement = owner.GetFirstChild<A.SchemeColor>();
            }

            if (colorElement == null) {
                colorElement = owner.GetFirstChild<A.SystemColor>();
            }

            if (colorElement == null) {
                return false;
            }

            solidFill.Append((OpenXmlElement)colorElement.CloneNode(true));
            OfficeColor? resolved = PowerPointThemeColorResolver.ResolveSolidFillOfficeColor(solidFill, colorScheme);
            if (!resolved.HasValue) {
                return false;
            }

            color = resolved.Value;
            return true;
        }
    }
}
