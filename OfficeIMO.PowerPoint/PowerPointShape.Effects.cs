using System;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.PowerPoint {
    public abstract partial class PowerPointShape {
        /// <summary>
        ///     Applies a drop shadow to the shape.
        /// </summary>
        public void SetShadow(string color, OfficeLength? blur = null, OfficeLength? distance = null,
            double angleDegrees = 45, int transparencyPercent = 35, bool rotateWithShape = false) {
            OfficeLength resolvedBlur = blur ?? OfficeLength.Points(4);
            OfficeLength resolvedDistance = distance ?? OfficeLength.Points(3);
            if (string.IsNullOrWhiteSpace(color)) {
                throw new ArgumentException("Shadow color cannot be null or empty.", nameof(color));
            }
            if (resolvedBlur < OfficeLength.Zero) {
                throw new ArgumentOutOfRangeException(nameof(blur));
            }
            if (resolvedDistance < OfficeLength.Zero) {
                throw new ArgumentOutOfRangeException(nameof(distance));
            }
            if (transparencyPercent is < 0 or > 100) {
                throw new ArgumentOutOfRangeException(nameof(transparencyPercent), "Transparency must be between 0 and 100.");
            }

            A.OuterShadow? shadow = GetOuterShadow(create: true);
            if (shadow == null) {
                return;
            }

            shadow.BlurRadius = resolvedBlur.ToEmus();
            shadow.Distance = resolvedDistance.ToEmus();
            shadow.Direction = ToShadowAngle(angleDegrees);
            shadow.RotateWithShape = rotateWithShape;

            RemoveShadowColors(shadow);
            A.RgbColorModelHex rgb = new() { Val = color };
            int alpha = 100000 - transparencyPercent * 1000;
            rgb.Append(new A.Alpha { Val = alpha });
            shadow.Append(rgb);
        }

        /// <summary>
        ///     Removes any drop shadow from the shape.
        /// </summary>
        public void ClearShadow() {
            A.EffectList? effects = GetEffectList();
            if (effects == null) {
                return;
            }

            effects.GetFirstChild<A.OuterShadow>()?.Remove();
            if (!effects.HasChildren) {
                effects.Remove();
            }
        }

        /// <summary>
        ///     Applies a glow effect to the shape.
        /// </summary>
        public void SetGlow(string color, OfficeLength? radius = null, int transparencyPercent = 30) {
            OfficeLength resolvedRadius = radius ?? OfficeLength.Points(4);
            if (string.IsNullOrWhiteSpace(color)) {
                throw new ArgumentException("Glow color cannot be null or empty.", nameof(color));
            }
            if (resolvedRadius < OfficeLength.Zero) {
                throw new ArgumentOutOfRangeException(nameof(radius));
            }
            if (transparencyPercent is < 0 or > 100) {
                throw new ArgumentOutOfRangeException(nameof(transparencyPercent), "Transparency must be between 0 and 100.");
            }

            A.Glow? glow = GetGlow(create: true);
            if (glow == null) {
                return;
            }

            glow.Radius = resolvedRadius.ToEmus();
            RemoveGlowColors(glow);

            A.RgbColorModelHex rgb = new() { Val = color };
            int alpha = 100000 - transparencyPercent * 1000;
            rgb.Append(new A.Alpha { Val = alpha });
            glow.Append(rgb);
        }

        /// <summary>
        ///     Removes any glow effect from the shape.
        /// </summary>
        public void ClearGlow() {
            A.EffectList? effects = GetEffectList();
            if (effects == null) {
                return;
            }

            effects.GetFirstChild<A.Glow>()?.Remove();
            if (!effects.HasChildren) {
                effects.Remove();
            }
        }

        /// <summary>
        ///     Applies a soft edges effect to the shape.
        /// </summary>
        public void SetSoftEdges(OfficeLength radius) {
            if (radius < OfficeLength.Zero) {
                throw new ArgumentOutOfRangeException(nameof(radius));
            }

            A.SoftEdge? softEdge = GetSoftEdge(create: true);
            if (softEdge == null) {
                return;
            }

            softEdge.Radius = radius.ToEmus();
        }

        /// <summary>
        ///     Removes any soft edges effect from the shape.
        /// </summary>
        public void ClearSoftEdges() {
            A.EffectList? effects = GetEffectList();
            if (effects == null) {
                return;
            }

            effects.GetFirstChild<A.SoftEdge>()?.Remove();
            if (!effects.HasChildren) {
                effects.Remove();
            }
        }

        /// <summary>
        ///     Applies a blur effect to the shape.
        /// </summary>
        public void SetBlur(OfficeLength radius, bool grow = false) {
            if (radius < OfficeLength.Zero) {
                throw new ArgumentOutOfRangeException(nameof(radius));
            }

            A.Blur? blur = GetBlur(create: true);
            if (blur == null) {
                return;
            }

            blur.Radius = radius.ToEmus();
            blur.Grow = grow;
        }

        /// <summary>
        ///     Removes any blur effect from the shape.
        /// </summary>
        public void ClearBlur() {
            A.EffectList? effects = GetEffectList();
            if (effects == null) {
                return;
            }

            effects.GetFirstChild<A.Blur>()?.Remove();
            if (!effects.HasChildren) {
                effects.Remove();
            }
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

            A.Reflection? reflection = GetReflection(create: true);
            if (reflection == null) {
                return;
            }

            reflection.BlurRadius = resolvedBlur.ToEmus();
            reflection.Distance = resolvedDistance.ToEmus();
            reflection.Direction = ToShadowAngle(directionDegrees);
            reflection.FadeDirection = ToShadowAngle(fadeDirectionDegrees);
            reflection.StartOpacity = ToAlphaValue(startOpacityPercent);
            reflection.EndAlpha = ToAlphaValue(endOpacityPercent);
            reflection.StartPosition = ToPercentValue(startPositionPercent);
            reflection.EndPosition = ToPercentValue(endPositionPercent);
            reflection.Alignment = (alignment ?? PowerPointRectangleAlignment.Bottom).ToOpenXml();
            reflection.RotateWithShape = rotateWithShape;
        }

        /// <summary>
        ///     Removes any reflection effect from the shape.
        /// </summary>
        public void ClearReflection() {
            A.EffectList? effects = GetEffectList();
            if (effects == null) {
                return;
            }

            effects.GetFirstChild<A.Reflection>()?.Remove();
            if (!effects.HasChildren) {
                effects.Remove();
            }
        }


        private A.EffectList? GetEffectList(bool create = false) {
            ShapeProperties? props = GetShapeProperties(create);
            if (props == null) {
                return null;
            }

            A.EffectList? effectList = props.GetFirstChild<A.EffectList>();
            if (effectList == null && create) {
                effectList = new A.EffectList();
                InsertShapePropertyChild(props, effectList);
            }

            return effectList;
        }

        private A.OuterShadow? GetOuterShadow(bool create = false) {
            A.EffectList? effects = GetEffectList(create);
            if (effects == null) {
                return null;
            }

            A.OuterShadow? shadow = effects.GetFirstChild<A.OuterShadow>();
            if (shadow == null && create) {
                shadow = new A.OuterShadow();
                InsertEffectChild(effects, shadow);
            }

            return shadow;
        }

        private A.Blur? GetBlur(bool create = false) {
            A.EffectList? effects = GetEffectList(create);
            if (effects == null) {
                return null;
            }

            A.Blur? blur = effects.GetFirstChild<A.Blur>();
            if (blur == null && create) {
                blur = new A.Blur();
                InsertEffectChild(effects, blur);
            }

            return blur;
        }

        private A.Reflection? GetReflection(bool create = false) {
            A.EffectList? effects = GetEffectList(create);
            if (effects == null) {
                return null;
            }

            A.Reflection? reflection = effects.GetFirstChild<A.Reflection>();
            if (reflection == null && create) {
                reflection = new A.Reflection();
                InsertEffectChild(effects, reflection);
            }

            return reflection;
        }

        private A.Glow? GetGlow(bool create = false) {
            A.EffectList? effects = GetEffectList(create);
            if (effects == null) {
                return null;
            }

            A.Glow? glow = effects.GetFirstChild<A.Glow>();
            if (glow == null && create) {
                glow = new A.Glow();
                InsertEffectChild(effects, glow);
            }

            return glow;
        }

        private A.SoftEdge? GetSoftEdge(bool create = false) {
            A.EffectList? effects = GetEffectList(create);
            if (effects == null) {
                return null;
            }

            A.SoftEdge? softEdge = effects.GetFirstChild<A.SoftEdge>();
            if (softEdge == null && create) {
                softEdge = new A.SoftEdge();
                InsertEffectChild(effects, softEdge);
            }

            return softEdge;
        }

        private static void InsertEffectChild(A.EffectList effects, OpenXmlElement child) {
            int childOrder = GetEffectChildOrder(child);
            OpenXmlElement? insertBefore = effects.ChildElements
                .FirstOrDefault(existing => GetEffectChildOrder(existing) > childOrder);

            if (insertBefore != null) {
                effects.InsertBefore(child, insertBefore);
            } else {
                effects.Append(child);
            }
        }

        private static int GetEffectChildOrder(OpenXmlElement child) {
            return child switch {
                A.Blur => 0,
                A.FillOverlay => 1,
                A.Glow => 2,
                A.InnerShadow => 3,
                A.OuterShadow => 4,
                A.PresetShadow => 5,
                A.Reflection => 6,
                A.SoftEdge => 7,
                _ => 100
            };
        }

        private static void RemoveShadowColors(A.OuterShadow shadow) {
            shadow.RemoveAllChildren<A.RgbColorModelHex>();
            shadow.RemoveAllChildren<A.SchemeColor>();
            shadow.RemoveAllChildren<A.SystemColor>();
            shadow.RemoveAllChildren<A.PresetColor>();
        }

        private static void RemoveGlowColors(A.Glow glow) {
            glow.RemoveAllChildren<A.RgbColorModelHex>();
            glow.RemoveAllChildren<A.SchemeColor>();
            glow.RemoveAllChildren<A.SystemColor>();
            glow.RemoveAllChildren<A.PresetColor>();
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
