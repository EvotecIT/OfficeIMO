using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.OpenXml.Internal;

namespace OfficeIMO.PowerPoint {
    public abstract partial class PowerPointShape {
        /// <summary>
        ///     Applies a drop shadow to the shape.
        /// </summary>
        public void SetShadow(OfficeColor color, OfficeLength? blur = null, OfficeLength? distance = null,
            double angleDegrees = 45, int transparencyPercent = 35, bool rotateWithShape = false) {
            OfficeLength resolvedBlur = blur ?? OfficeLength.Points(4);
            OfficeLength resolvedDistance = distance ?? OfficeLength.Points(3);
            if (resolvedBlur < OfficeLength.Zero) {
                throw new ArgumentOutOfRangeException(nameof(blur));
            }
            if (resolvedDistance < OfficeLength.Zero) {
                throw new ArgumentOutOfRangeException(nameof(distance));
            }
            if (transparencyPercent is < 0 or > 100) {
                throw new ArgumentOutOfRangeException(nameof(transparencyPercent), "Transparency must be between 0 and 100.");
            }

            OfficeTheme? theme = PowerPointDrawingColor.NeedsTheme(color) ? PowerPointDrawingColor.ThemeFor(OwnerPartRef, OwnerSlide) : null;
            OfficeXmlNode shadow = GetEffect("outerShdw", create: true);
            if (shadow.IsNone) {
                return;
            }

            shadow.SetInt64Attribute("blurRad", resolvedBlur.ToEmus());
            shadow.SetInt64Attribute("dist", resolvedDistance.ToEmus());
            shadow.SetInt32Attribute("dir", ToShadowAngle(angleDegrees));
            shadow.SetBooleanAttribute("rotWithShape", rotateWithShape);

            RemoveEffectColors(shadow);
            AppendEffectColor(shadow, color, theme, transparencyPercent);
            OwnerPartRef.MarkEdited();
        }

        /// <summary>
        ///     Removes any drop shadow from the shape.
        /// </summary>
        public void ClearShadow() {
            RemoveEffect("outerShdw");
        }

        /// <summary>
        ///     Applies a glow effect to the shape.
        /// </summary>
        public void SetGlow(OfficeColor color, OfficeLength? radius = null, int transparencyPercent = 30) {
            OfficeLength resolvedRadius = radius ?? OfficeLength.Points(4);
            if (resolvedRadius < OfficeLength.Zero) {
                throw new ArgumentOutOfRangeException(nameof(radius));
            }
            if (transparencyPercent is < 0 or > 100) {
                throw new ArgumentOutOfRangeException(nameof(transparencyPercent), "Transparency must be between 0 and 100.");
            }

            OfficeTheme? theme = PowerPointDrawingColor.NeedsTheme(color) ? PowerPointDrawingColor.ThemeFor(OwnerPartRef, OwnerSlide) : null;
            OfficeXmlNode glow = GetEffect("glow", create: true);
            if (glow.IsNone) {
                return;
            }

            glow.SetInt64Attribute("rad", resolvedRadius.ToEmus());
            RemoveEffectColors(glow);
            AppendEffectColor(glow, color, theme, transparencyPercent);
            OwnerPartRef.MarkEdited();
        }

        /// <summary>
        ///     Removes any glow effect from the shape.
        /// </summary>
        public void ClearGlow() {
            RemoveEffect("glow");
        }

        /// <summary>
        ///     Applies a soft edges effect to the shape.
        /// </summary>
        public void SetSoftEdges(OfficeLength radius) {
            if (radius < OfficeLength.Zero) {
                throw new ArgumentOutOfRangeException(nameof(radius));
            }

            OfficeXmlNode softEdge = GetEffect("softEdge", create: true);
            if (softEdge.IsNone) {
                return;
            }

            softEdge.SetInt64Attribute("rad", radius.ToEmus());
            OwnerPartRef.MarkEdited();
        }

        /// <summary>
        ///     Removes any soft edges effect from the shape.
        /// </summary>
        public void ClearSoftEdges() {
            RemoveEffect("softEdge");
        }

        /// <summary>
        ///     Applies a blur effect to the shape.
        /// </summary>
        public void SetBlur(OfficeLength radius, bool grow = false) {
            if (radius < OfficeLength.Zero) {
                throw new ArgumentOutOfRangeException(nameof(radius));
            }

            OfficeXmlNode blur = GetEffect("blur", create: true);
            if (blur.IsNone) {
                return;
            }

            blur.SetInt64Attribute("rad", radius.ToEmus());
            blur.SetBooleanAttribute("grow", grow);
            OwnerPartRef.MarkEdited();
        }

        /// <summary>
        ///     Removes any blur effect from the shape.
        /// </summary>
        public void ClearBlur() {
            RemoveEffect("blur");
        }

        /// <summary>
        ///     Applies a reflection effect to the shape.
        /// </summary>
        public void SetReflection(
            OfficeLength? blur = null,
            OfficeLength? distance = null,
            double directionDegrees = 270,
            double fadeDirectionDegrees = 90,
            int startOpacityPercent = 50,
            int endOpacityPercent = 0,
            int startPositionPercent = 0,
            int endPositionPercent = 100,
            PowerPointRectangleAlignment? alignment = null,
            bool rotateWithShape = false) {
            OfficeLength resolvedBlur = blur ?? OfficeLength.Points(4);
            OfficeLength resolvedDistance = distance ?? OfficeLength.Points(2);
            if (resolvedBlur < OfficeLength.Zero) {
                throw new ArgumentOutOfRangeException(nameof(blur));
            }
            if (resolvedDistance < OfficeLength.Zero) {
                throw new ArgumentOutOfRangeException(nameof(distance));
            }
            if (startOpacityPercent is < 0 or > 100) {
                throw new ArgumentOutOfRangeException(nameof(startOpacityPercent), "Opacity must be between 0 and 100.");
            }
            if (endOpacityPercent is < 0 or > 100) {
                throw new ArgumentOutOfRangeException(nameof(endOpacityPercent), "Opacity must be between 0 and 100.");
            }
            if (startPositionPercent is < 0 or > 100) {
                throw new ArgumentOutOfRangeException(nameof(startPositionPercent), "Position must be between 0 and 100.");
            }
            if (endPositionPercent is < 0 or > 100) {
                throw new ArgumentOutOfRangeException(nameof(endPositionPercent), "Position must be between 0 and 100.");
            }

            OfficeXmlNode reflection = GetEffect("reflection", create: true);
            if (reflection.IsNone) {
                return;
            }

            // Schema order, as the SDK writes the attributes of a new element.
            reflection.SetInt64Attribute("blurRad", resolvedBlur.ToEmus());
            reflection.SetInt32Attribute("stA", ToAlphaValue(startOpacityPercent));
            reflection.SetInt32Attribute("stPos", ToPercentValue(startPositionPercent));
            reflection.SetInt32Attribute("endA", ToAlphaValue(endOpacityPercent));
            reflection.SetInt32Attribute("endPos", ToPercentValue(endPositionPercent));
            reflection.SetInt64Attribute("dist", resolvedDistance.ToEmus());
            reflection.SetInt32Attribute("dir", ToShadowAngle(directionDegrees));
            reflection.SetInt32Attribute("fadeDir", ToShadowAngle(fadeDirectionDegrees));
            reflection.SetAttribute("algn", (alignment ?? PowerPointRectangleAlignment.Bottom).ToOpenXmlToken());
            reflection.SetBooleanAttribute("rotWithShape", rotateWithShape);
            OwnerPartRef.MarkEdited();
        }

        /// <summary>
        ///     Removes any reflection effect from the shape.
        /// </summary>
        public void ClearReflection() {
            RemoveEffect("reflection");
        }


        // The children of a:effectLst in sequence order.
        private static readonly OfficeXmlChildOrder EffectOrder = new(
            new[] { D, D, D, D, D, D, D, D },
            new[] { "blur", "fillOverlay", "glow", "innerShdw", "outerShdw", "prstShdw", "reflection", "softEdge" });

        private OfficeXmlNode GetEffectList(bool create = false) {
            OfficeXmlNode properties = ShapePropertiesElement(create);
            if (properties.IsNone) return default;
            OfficeXmlNode effects = properties.Child(D, "effectLst");
            if (effects.IsNone && create) effects = properties.InsertOrdered(properties.CreateElement(D, "effectLst"), ShapePropertyOrder);
            return effects;
        }

        private OfficeXmlNode GetEffect(string localName, bool create = false) {
            OfficeXmlNode effects = GetEffectList(create);
            if (effects.IsNone) return default;
            OfficeXmlNode effect = effects.Child(D, localName);
            if (effect.IsNone && create) effect = effects.InsertOrdered(effects.CreateElement(D, localName), EffectOrder);
            return effect;
        }

        // Removes one effect, and the effect list when it is left empty.
        private void RemoveEffect(string localName) {
            OfficeXmlNode effects = GetEffectList();
            if (effects.IsNone) return;
            OfficeXmlNode effect = effects.Child(D, localName);
            if (!effect.IsNone) effect.Remove();
            if (effects.FirstChild.IsNone) effects.Remove();
            OwnerPartRef.MarkEdited();
        }

        private static void RemoveEffectColors(OfficeXmlNode effect) {
            OfficeXmlNode child = effect.FirstChild;
            while (!child.IsNone) {
                OfficeXmlNode next = child.NextSibling;
                if (child.NamespaceId == D && child.LocalName is "srgbClr" or "schemeClr" or "sysClr" or "prstClr") child.Remove();
                child = next;
            }
        }

        // Appends the effect's color with the transparency applied on top of the color's own opacity.
        private static void AppendEffectColor(OfficeXmlNode effect, OfficeColor color, OfficeTheme? theme, int transparencyPercent) {
            OfficeXmlNode element = PowerPointDrawingColor.Append(effect, color, theme);
            int alpha = element.Child(D, "alpha").Int32Attribute("val") ?? 100000;
            element.RemoveChildren(D, "alpha");
            element.Append(element.CreateElement(D, "alpha")).SetInt32Attribute("val", (int)Math.Round(alpha * (100 - transparencyPercent) / 100D));
        }

        private static int ToShadowAngle(double degrees) {
            double normalized = degrees % 360d;
            if (normalized < 0) {
                normalized += 360d;
            }
            return (int)Math.Round(normalized * 60000d);
        }

        private static int ToAlphaValue(int percent) {
            return percent * 1000;
        }

        private static int ToPercentValue(int percent) {
            return percent * 1000;
        }
    }
}
