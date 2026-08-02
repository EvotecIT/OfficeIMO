using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.PowerPoint {
    public abstract partial class PowerPointShape {
        /// <summary>
        ///     Gets or sets the fill color of the shape in hex format (e.g. "FF0000").
        /// </summary>
        public string? FillColor {
            get {
                ShapeProperties? props = GetShapeProperties();
                A.SolidFill? solid = props?.GetFirstChild<A.SolidFill>();
                return solid?.RgbColorModelHex?.Val;
            }
            set {
                ShapeProperties? props = GetShapeProperties(create: value != null);
                if (props == null) {
                    return;
                }

                if (value == null) {
                    props.RemoveAllChildren<A.SolidFill>();
                    return;
                }
                RemoveFillChoiceChildren(props);
                InsertShapePropertyChild(props,
                    new A.SolidFill(new A.RgbColorModelHex { Val = value }));
            }
        }

        /// <summary>
        ///     Gets or sets the fill transparency percentage (0-100). 0 = opaque, 100 = fully transparent.
        /// </summary>
        public int? FillTransparency {
            get {
                ShapeProperties? props = GetShapeProperties();
                IReadOnlyList<OpenXmlCompositeElement> localColors = props == null
                    ? Array.Empty<OpenXmlCompositeElement>()
                    : GetFillColorChoices(props);
                if (props != null && HasExplicitFillChoice(props)) {
                    A.AlphaModulationFixed? imageAlpha = props
                        .GetFirstChild<A.BlipFill>()?.Blip?
                        .GetFirstChild<A.AlphaModulationFixed>();
                    if (imageAlpha?.Amount?.Value is int imageAmount) {
                        return (int)Math.Round((100000 - imageAmount)
                            / 1000D);
                    }
                    if (localColors.Count == 0) return null;
                    int? localAlpha = localColors[0]
                        .GetFirstChild<A.Alpha>()?.Val?.Value;
                    if (localAlpha == null || localColors.Any(color =>
                            color.GetFirstChild<A.Alpha>()?.Val?.Value
                                != localAlpha)) {
                        return null;
                    }
                    return (int)Math.Round((100000 - localAlpha.Value)
                        / 1000D);
                }
                OpenXmlCompositeElement? color =
                    GetShapeStyleFillColorChoice(createPlaceholder: false);
                int? alpha = color?.GetFirstChild<A.Alpha>()?.Val?.Value;
                return alpha == null
                    ? null
                    : (int)Math.Round((100000 - alpha.Value) / 1000D);
            }
            set {
                if (value is < 0 or > 100) {
                    throw new ArgumentOutOfRangeException(nameof(value), "Transparency must be between 0 and 100.");
                }
                SetFillOpacity(value == null ? null : 1D - value.Value / 100D);
            }
        }

        internal void SetFillOpacity(double? opacity) {
            if (opacity.HasValue &&
                (double.IsNaN(opacity.Value) || double.IsInfinity(opacity.Value)
                    || opacity.Value < 0D || opacity.Value > 1D)) {
                throw new ArgumentOutOfRangeException(nameof(opacity),
                    "Opacity must be between 0 and 1.");
            }

            ShapeProperties? props = GetShapeProperties(create: opacity != null);
            if (props == null) return;
            A.SolidFill? solid = props.GetFirstChild<A.SolidFill>();
            if (solid == null) {
                if (HasExplicitFillChoice(props)) {
                    A.BlipFill? pictureFill = props
                        .GetFirstChild<A.BlipFill>();
                    if (pictureFill != null) {
                        A.Blip blip = pictureFill.Blip
                            ?? throw new InvalidOperationException(
                                "The picture fill has no image reference to receive transparency.");
                        SetBlipAlpha(blip, opacity);
                        return;
                    }
                    foreach (OpenXmlCompositeElement fillColor in
                             GetFillColorChoices(props)) {
                        SetFillColorAlpha(fillColor, opacity);
                    }
                    return;
                }
                OpenXmlElement? themeFill = ResolveThemeFill();
                if (opacity != null && themeFill != null
                    && !themeFill.Descendants<A.SchemeColor>().Any(color =>
                        color.Val?.Value == A.SchemeColorValues.PhColor)) {
                    OpenXmlElement materialized = themeFill.CloneNode(true);
                    RemoveFillChoiceChildren(props);
                    InsertShapePropertyChild(props, materialized);
                    foreach (OpenXmlCompositeElement fillColor in
                             GetFillColorChoices(props)) {
                        SetFillColorAlpha(fillColor, opacity);
                    }
                    return;
                }
                OpenXmlCompositeElement? styleColor =
                    GetShapeStyleFillColorChoice(createPlaceholder: opacity != null);
                if (styleColor != null) {
                    SetFillColorAlpha(styleColor, opacity);
                    return;
                }
                if (opacity == null) return;
                solid = new A.SolidFill(new A.RgbColorModelHex { Val = "FFFFFF" });
                InsertShapePropertyChild(props, solid);
            }

            OpenXmlCompositeElement? color = GetColorChoice(solid);
            if (opacity == null) {
                color?.GetFirstChild<A.Alpha>()?.Remove();
                return;
            }
            if (color == null) {
                color = new A.RgbColorModelHex { Val = "FFFFFF" };
                solid.Append(color);
            }

            SetFillColorAlpha(color, opacity);
        }

        private static void SetBlipAlpha(A.Blip blip, double? opacity) {
            A.AlphaModulationFixed? alpha = blip
                .GetFirstChild<A.AlphaModulationFixed>();
            if (!opacity.HasValue) {
                alpha?.Remove();
                return;
            }
            alpha ??= new A.AlphaModulationFixed();
            alpha.Amount = checked((int)Math.Round(opacity.Value * 100000D,
                MidpointRounding.AwayFromZero));
            if (alpha.Parent == null) blip.Append(alpha);
        }

        private OpenXmlElement? ResolveThemeFill() {
            A.FillReference? reference = Element switch {
                Shape shape => shape.ShapeStyle?.FillReference,
                ConnectionShape connector => connector.ShapeStyle?.FillReference,
                _ => null
            };
            uint? index = reference?.Index?.Value;
            if (OwnerSlide == null || !index.HasValue) return null;
            A.FormatScheme? formatScheme = OwnerSlide.SlidePart
                .ThemeOverridePart?.ThemeOverride?.FormatScheme
                ?? OwnerSlide.SlidePart.SlideLayoutPart?.ThemeOverridePart?
                    .ThemeOverride?.FormatScheme
                ?? OwnerSlide.SlidePart.SlideLayoutPart?.SlideMasterPart?
                    .ThemePart?.Theme?.ThemeElements?.FormatScheme;
            if (formatScheme == null) return null;
            OpenXmlElementList fills;
            uint zeroBased;
            if (index.Value >= 1001U) {
                fills = formatScheme.GetFirstChild<A.BackgroundFillStyleList>()?
                    .ChildElements ?? default;
                zeroBased = index.Value - 1001U;
            } else {
                if (index.Value < 1U) return null;
                fills = formatScheme.GetFirstChild<A.FillStyleList>()?
                    .ChildElements ?? default;
                zeroBased = index.Value - 1U;
            }
            return zeroBased < unchecked((uint)fills.Count)
                ? fills[unchecked((int)zeroBased)]
                : null;
        }

        private OpenXmlCompositeElement? GetShapeStyleFillColorChoice(
            bool createPlaceholder) {
            A.FillReference? reference = Element switch {
                Shape shape => shape.ShapeStyle?.FillReference,
                ConnectionShape connector => connector.ShapeStyle?.FillReference,
                _ => null
            };
            if (reference == null) return null;

            OpenXmlCompositeElement? color = GetColorChoice(reference);
            if (color == null && createPlaceholder) {
                color = new A.SchemeColor { Val = A.SchemeColorValues.PhColor };
                reference.Append(color);
            }
            return color;
        }

        private static void SetFillColorAlpha(OpenXmlCompositeElement color,
            double? opacity) {
            if (!opacity.HasValue) {
                color.GetFirstChild<A.Alpha>()?.Remove();
                return;
            }

            A.Alpha? alpha = color.GetFirstChild<A.Alpha>() ?? new A.Alpha();
            alpha.Val = checked((int)Math.Round(opacity.Value * 100000D,
                MidpointRounding.AwayFromZero));
            if (alpha.Parent == null) color.Append(alpha);
        }

        private static OpenXmlCompositeElement? GetColorChoice(
            OpenXmlCompositeElement parent) => parent.ChildElements
                .OfType<OpenXmlCompositeElement>()
                .FirstOrDefault(element =>
                    element is A.RgbColorModelPercentage
                    || element is A.RgbColorModelHex
                    || element is A.HslColor
                    || element is A.SystemColor
                    || element is A.SchemeColor
                    || element is A.PresetColor);

        private static bool HasExplicitFillChoice(
            OpenXmlCompositeElement parent) =>
            parent.ChildElements.Any(child => child is A.NoFill
                or A.SolidFill or A.GradientFill or A.BlipFill
                or A.PatternFill or A.GroupFill);

        private static void RemoveFillChoiceChildren(
            OpenXmlCompositeElement parent) {
            parent.RemoveAllChildren<A.NoFill>();
            parent.RemoveAllChildren<A.SolidFill>();
            parent.RemoveAllChildren<A.GradientFill>();
            parent.RemoveAllChildren<A.BlipFill>();
            parent.RemoveAllChildren<A.PatternFill>();
            parent.RemoveAllChildren<A.GroupFill>();
        }

        private static IReadOnlyList<OpenXmlCompositeElement>
            GetFillColorChoices(OpenXmlCompositeElement parent) {
            A.SolidFill? solid = parent.GetFirstChild<A.SolidFill>();
            if (solid != null) {
                OpenXmlCompositeElement? color = GetColorChoice(solid);
                return color == null
                    ? Array.Empty<OpenXmlCompositeElement>()
                    : new[] { color };
            }

            A.GradientFill? gradient = parent.GetFirstChild<A.GradientFill>();
            if (gradient != null) {
                return gradient.Descendants<A.GradientStop>()
                    .Select(GetColorChoice)
                    .Where(color => color != null)
                    .Cast<OpenXmlCompositeElement>()
                    .ToArray();
            }

            A.PatternFill? pattern = parent.GetFirstChild<A.PatternFill>();
            if (pattern != null) {
                return pattern.ChildElements
                    .OfType<OpenXmlCompositeElement>()
                    .Select(GetColorChoice)
                    .Where(color => color != null)
                    .Cast<OpenXmlCompositeElement>()
                    .ToArray();
            }

            return Array.Empty<OpenXmlCompositeElement>();
        }

        /// <summary>
        ///     Gets or sets rotation in degrees.
        /// </summary>
        public double? Rotation {
            get {
                int? rotation = GetRotation();
                return rotation != null ? rotation.Value / 60000d : null;
            }
            set {
                int? rotation = value != null ? (int)Math.Round(value.Value * 60000d) : null;
                SetRotation(rotation);
            }
        }

        /// <summary>
        ///     Gets or sets horizontal flip.
        /// </summary>
        public bool? HorizontalFlip {
            get => GetHorizontalFlip();
            set => SetHorizontalFlip(value);
        }

        /// <summary>
        ///     Gets or sets vertical flip.
        /// </summary>
        public bool? VerticalFlip {
            get => GetVerticalFlip();
            set => SetVerticalFlip(value);
        }

        private int? GetRotation() {
            return Element switch {
                Shape s => s.ShapeProperties?.Transform2D?.Rotation?.Value,
                ConnectionShape c => c.ShapeProperties?.Transform2D?.Rotation?.Value,
                Picture p => p.ShapeProperties?.Transform2D?.Rotation?.Value,
                GraphicFrame g => g.Transform?.Rotation?.Value,
                GroupShape g => g.GroupShapeProperties?.TransformGroup?.Rotation?.Value,
                _ => null
            };
        }

        private void SetRotation(int? rotation) {
            switch (Element) {
                case Shape s:
                    s.ShapeProperties ??= new ShapeProperties();
                    s.ShapeProperties.Transform2D ??= new A.Transform2D();
                    s.ShapeProperties.Transform2D.Rotation = rotation;
                    break;
                case ConnectionShape c:
                    c.ShapeProperties ??= new ShapeProperties();
                    c.ShapeProperties.Transform2D ??= new A.Transform2D();
                    c.ShapeProperties.Transform2D.Rotation = rotation;
                    break;
                case Picture p:
                    p.ShapeProperties ??= new ShapeProperties();
                    p.ShapeProperties.Transform2D ??= new A.Transform2D();
                    p.ShapeProperties.Transform2D.Rotation = rotation;
                    break;
                case GraphicFrame g:
                    g.Transform ??= new Transform();
                    g.Transform.Rotation = rotation;
                    break;
                case GroupShape g:
                    EnsureTransformGroup(g).Rotation = rotation;
                    break;
            }
        }

        private bool? GetHorizontalFlip() {
            return Element switch {
                Shape s => s.ShapeProperties?.Transform2D?.HorizontalFlip?.Value,
                ConnectionShape c => c.ShapeProperties?.Transform2D?.HorizontalFlip?.Value,
                Picture p => p.ShapeProperties?.Transform2D?.HorizontalFlip?.Value,
                GraphicFrame g => g.Transform?.HorizontalFlip?.Value,
                GroupShape g => g.GroupShapeProperties?.TransformGroup?.HorizontalFlip?.Value,
                _ => null
            };
        }

        private void SetHorizontalFlip(bool? value) {
            switch (Element) {
                case Shape s:
                    s.ShapeProperties ??= new ShapeProperties();
                    s.ShapeProperties.Transform2D ??= new A.Transform2D();
                    s.ShapeProperties.Transform2D.HorizontalFlip = value;
                    break;
                case ConnectionShape c:
                    c.ShapeProperties ??= new ShapeProperties();
                    c.ShapeProperties.Transform2D ??= new A.Transform2D();
                    c.ShapeProperties.Transform2D.HorizontalFlip = value;
                    break;
                case Picture p:
                    p.ShapeProperties ??= new ShapeProperties();
                    p.ShapeProperties.Transform2D ??= new A.Transform2D();
                    p.ShapeProperties.Transform2D.HorizontalFlip = value;
                    break;
                case GraphicFrame g:
                    g.Transform ??= new Transform();
                    g.Transform.HorizontalFlip = value;
                    break;
                case GroupShape g:
                    EnsureTransformGroup(g).HorizontalFlip = value;
                    break;
            }
        }

        private bool? GetVerticalFlip() {
            return Element switch {
                Shape s => s.ShapeProperties?.Transform2D?.VerticalFlip?.Value,
                ConnectionShape c => c.ShapeProperties?.Transform2D?.VerticalFlip?.Value,
                Picture p => p.ShapeProperties?.Transform2D?.VerticalFlip?.Value,
                GraphicFrame g => g.Transform?.VerticalFlip?.Value,
                GroupShape g => g.GroupShapeProperties?.TransformGroup?.VerticalFlip?.Value,
                _ => null
            };
        }

        private void SetVerticalFlip(bool? value) {
            switch (Element) {
                case Shape s:
                    s.ShapeProperties ??= new ShapeProperties();
                    s.ShapeProperties.Transform2D ??= new A.Transform2D();
                    s.ShapeProperties.Transform2D.VerticalFlip = value;
                    break;
                case ConnectionShape c:
                    c.ShapeProperties ??= new ShapeProperties();
                    c.ShapeProperties.Transform2D ??= new A.Transform2D();
                    c.ShapeProperties.Transform2D.VerticalFlip = value;
                    break;
                case Picture p:
                    p.ShapeProperties ??= new ShapeProperties();
                    p.ShapeProperties.Transform2D ??= new A.Transform2D();
                    p.ShapeProperties.Transform2D.VerticalFlip = value;
                    break;
                case GraphicFrame g:
                    g.Transform ??= new Transform();
                    g.Transform.VerticalFlip = value;
                    break;
                case GroupShape g:
                    EnsureTransformGroup(g).VerticalFlip = value;
                    break;
            }
        }
    }
}
