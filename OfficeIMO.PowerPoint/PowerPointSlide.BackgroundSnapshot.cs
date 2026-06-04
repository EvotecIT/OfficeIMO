using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.PowerPoint {
    public partial class PowerPointSlide {
        /// <summary>
        /// Returns a detached snapshot of the slide background fill that exporters can consume without Open XML coupling.
        /// </summary>
        public PowerPointSlideBackground GetBackground() {
            (BackgroundProperties? properties, OpenXmlPart? ownerPart) = GetResolvedBackgroundProperties();
            if (properties == null || !properties.HasChildren) {
                (BackgroundStyleReference? styleReference, OpenXmlPart? styleOwnerPart) = GetResolvedBackgroundStyleReference();
                return styleReference != null
                    ? GetBackgroundStyleReference(styleReference, styleOwnerPart ?? _slidePart)
                    : PowerPointSlideBackground.None();
            }

            A.SolidFill? solidFill = properties.GetFirstChild<A.SolidFill>();
            string? solidColor = null;
            if (solidFill != null) {
                A.ColorScheme? colorScheme = GetThemePart(ownerPart ?? _slidePart)?.Theme?.ThemeElements?.ColorScheme;
                solidColor = ResolveSolidFillColor(solidFill, colorScheme, placeholderColor: null);
            }

            if (!string.IsNullOrWhiteSpace(solidColor)) {
                return PowerPointSlideBackground.SolidColor(solidColor!);
            }

            A.BlipFill? blipFill = properties.GetFirstChild<A.BlipFill>();
            if (blipFill != null) {
                return GetBackgroundImage(blipFill, ownerPart ?? _slidePart);
            }

            A.GradientFill? gradientFill = properties.GetFirstChild<A.GradientFill>();
            if (gradientFill != null) {
                A.ColorScheme? colorScheme = GetThemePart(ownerPart ?? _slidePart)?.Theme?.ThemeElements?.ColorScheme;
                return GetBackgroundGradient(gradientFill, colorScheme, placeholderColor: null);
            }

            return PowerPointSlideBackground.Unsupported("The slide background fill type is not currently supported by OfficeIMO exporters.");
        }

        private (BackgroundProperties? Properties, OpenXmlPart? OwnerPart) GetResolvedBackgroundProperties() {
            BackgroundProperties? slideProperties = SlideRoot.CommonSlideData?.Background?.BackgroundProperties;
            if (slideProperties != null && slideProperties.HasChildren) {
                return (slideProperties, _slidePart);
            }

            SlideLayoutPart? layoutPart = _slidePart.SlideLayoutPart;
            BackgroundProperties? layoutProperties = layoutPart?.SlideLayout?.CommonSlideData?.Background?.BackgroundProperties;
            if (layoutProperties != null && layoutProperties.HasChildren) {
                return (layoutProperties, layoutPart);
            }

            SlideMasterPart? masterPart = layoutPart?.SlideMasterPart;
            BackgroundProperties? masterProperties = masterPart?.SlideMaster?.CommonSlideData?.Background?.BackgroundProperties;
            if (masterProperties != null && masterProperties.HasChildren) {
                return (masterProperties, masterPart);
            }

            return (null, null);
        }

        private (BackgroundStyleReference? StyleReference, OpenXmlPart? OwnerPart) GetResolvedBackgroundStyleReference() {
            BackgroundStyleReference? slideReference = SlideRoot.CommonSlideData?.Background?.BackgroundStyleReference;
            if (slideReference != null) {
                return (slideReference, _slidePart);
            }

            SlideLayoutPart? layoutPart = _slidePart.SlideLayoutPart;
            BackgroundStyleReference? layoutReference = layoutPart?.SlideLayout?.CommonSlideData?.Background?.BackgroundStyleReference;
            if (layoutReference != null) {
                return (layoutReference, layoutPart);
            }

            SlideMasterPart? masterPart = layoutPart?.SlideMasterPart;
            BackgroundStyleReference? masterReference = masterPart?.SlideMaster?.CommonSlideData?.Background?.BackgroundStyleReference;
            if (masterReference != null) {
                return (masterReference, masterPart);
            }

            return (null, null);
        }

        private static PowerPointSlideBackground GetBackgroundStyleReference(BackgroundStyleReference styleReference, OpenXmlPart ownerPart) {
            ThemePart? themePart = GetThemePart(ownerPart);
            A.FormatScheme? formatScheme = themePart?.Theme?.ThemeElements?.FormatScheme;
            if (formatScheme == null) {
                return PowerPointSlideBackground.Unsupported("The slide background references a theme background style but no theme format scheme could be resolved.");
            }

            OpenXmlElement? fill = ResolveThemeBackgroundFill(formatScheme, styleReference.Index?.Value);
            if (fill == null) {
                return PowerPointSlideBackground.Unsupported("The slide background references a theme background style that could not be resolved.");
            }

            A.ColorScheme? colorScheme = themePart?.Theme?.ThemeElements?.ColorScheme;
            string? solidColor = ResolveSolidFillColor(fill as A.SolidFill ?? fill.GetFirstChild<A.SolidFill>(), colorScheme, styleReference.GetFirstChild<A.SchemeColor>());
            if (!string.IsNullOrWhiteSpace(solidColor)) {
                return PowerPointSlideBackground.SolidColor(solidColor!);
            }

            return PowerPointSlideBackground.Unsupported("The slide background references a theme background fill type that is not currently supported by OfficeIMO exporters.");
        }

        private static ThemePart? GetThemePart(OpenXmlPart ownerPart) {
            if (ownerPart is SlidePart slidePart) {
                return slidePart.SlideLayoutPart?.SlideMasterPart?.ThemePart;
            }

            if (ownerPart is SlideLayoutPart layoutPart) {
                return layoutPart.SlideMasterPart?.ThemePart;
            }

            return (ownerPart as SlideMasterPart)?.ThemePart;
        }

        private static OpenXmlElement? ResolveThemeBackgroundFill(A.FormatScheme formatScheme, uint? index) {
            if (!index.HasValue) {
                return null;
            }

            if (index.Value >= 1001U) {
                int backgroundIndex = (int)(index.Value - 1001U);
                return formatScheme.GetFirstChild<A.BackgroundFillStyleList>()?.ChildElements.ElementAtOrDefault(backgroundIndex);
            }

            if (index.Value >= 1U) {
                return formatScheme.GetFirstChild<A.FillStyleList>()?.ChildElements.ElementAtOrDefault((int)index.Value - 1);
            }

            return null;
        }

        private static string? ResolveSolidFillColor(A.SolidFill? solidFill, A.ColorScheme? colorScheme, A.SchemeColor? placeholderColor) {
            if (solidFill == null) {
                return null;
            }

            string? rgbColor = solidFill.RgbColorModelHex?.Val?.Value;
            if (!string.IsNullOrWhiteSpace(rgbColor)) {
                return ApplyColorTransforms(rgbColor, solidFill.RgbColorModelHex);
            }

            A.SchemeColor? schemeColor = solidFill.GetFirstChild<A.SchemeColor>();
            string? scheme = GetSchemeColorValue(schemeColor);
            if (IsPlaceholderSchemeColor(scheme)) {
                string? placeholderScheme = GetSchemeColorValue(placeholderColor);
                string? placeholderResolvedColor = ResolveSchemeColor(colorScheme, placeholderScheme);
                placeholderResolvedColor = ApplyColorTransforms(placeholderResolvedColor, placeholderColor);
                return ApplyColorTransforms(placeholderResolvedColor, schemeColor);
            }

            return ApplyColorTransforms(ResolveSchemeColor(colorScheme, scheme), schemeColor);
        }

        private static string? GetSchemeColorValue(A.SchemeColor? schemeColor) {
            string? attribute = schemeColor?.GetAttribute("val", string.Empty).Value;
            return !string.IsNullOrWhiteSpace(attribute)
                ? attribute
                : schemeColor?.Val?.Value.ToString();
        }

        private static bool IsPlaceholderSchemeColor(string? scheme) {
            return string.Equals(scheme, "Placeholder", StringComparison.OrdinalIgnoreCase)
                || string.Equals(scheme, "PlaceholderColor", StringComparison.OrdinalIgnoreCase)
                || string.Equals(scheme, "phClr", StringComparison.OrdinalIgnoreCase);
        }

        private static string? ResolveSchemeColor(A.ColorScheme? colorScheme, string? scheme) {
            if (colorScheme == null || string.IsNullOrWhiteSpace(scheme)) {
                return null;
            }

            OpenXmlCompositeElement? colorElement = scheme switch {
                "Dark1" or "dk1" or "Text1" or "tx1" => colorScheme.GetFirstChild<A.Dark1Color>(),
                "Light1" or "lt1" or "Background1" or "bg1" => colorScheme.GetFirstChild<A.Light1Color>(),
                "Dark2" or "dk2" or "Text2" or "tx2" => colorScheme.GetFirstChild<A.Dark2Color>(),
                "Light2" or "lt2" or "Background2" or "bg2" => colorScheme.GetFirstChild<A.Light2Color>(),
                "Accent1" or "accent1" => colorScheme.GetFirstChild<A.Accent1Color>(),
                "Accent2" or "accent2" => colorScheme.GetFirstChild<A.Accent2Color>(),
                "Accent3" or "accent3" => colorScheme.GetFirstChild<A.Accent3Color>(),
                "Accent4" or "accent4" => colorScheme.GetFirstChild<A.Accent4Color>(),
                "Accent5" or "accent5" => colorScheme.GetFirstChild<A.Accent5Color>(),
                "Accent6" or "accent6" => colorScheme.GetFirstChild<A.Accent6Color>(),
                "Hyperlink" or "hlink" => colorScheme.GetFirstChild<A.Hyperlink>(),
                "FollowedHyperlink" or "folHlink" => colorScheme.GetFirstChild<A.FollowedHyperlinkColor>(),
                _ => null
            };

            return colorElement?.GetFirstChild<A.RgbColorModelHex>()?.Val?.Value
                ?? colorElement?.GetFirstChild<A.SystemColor>()?.LastColor?.Value;
        }

        private static string? ApplyColorTransforms(string? hexColor, OpenXmlElement? colorElement) {
            if (string.IsNullOrWhiteSpace(hexColor) || colorElement == null) {
                return hexColor;
            }

            string color = hexColor!;
            if (color.Length != 6) {
                return hexColor;
            }

            if (!TryParseHexColor(color, out int red, out int green, out int blue)) {
                return hexColor;
            }

            foreach (OpenXmlElement transform in colorElement.ChildElements) {
                int? rawValue = GetTransformValue(transform);
                if (!rawValue.HasValue) {
                    continue;
                }

                double amount = Math.Max(0D, Math.Min(100000D, rawValue.Value)) / 100000D;
                switch (transform.LocalName) {
                    case "lumMod":
                        red = ClampColor(red * amount);
                        green = ClampColor(green * amount);
                        blue = ClampColor(blue * amount);
                        break;
                    case "lumOff":
                        red = ClampColor(red + 255D * amount);
                        green = ClampColor(green + 255D * amount);
                        blue = ClampColor(blue + 255D * amount);
                        break;
                    case "tint":
                        red = ClampColor(red + (255D - red) * amount);
                        green = ClampColor(green + (255D - green) * amount);
                        blue = ClampColor(blue + (255D - blue) * amount);
                        break;
                    case "shade":
                        red = ClampColor(red * amount);
                        green = ClampColor(green * amount);
                        blue = ClampColor(blue * amount);
                        break;
                }
            }

            return red.ToString("X2") + green.ToString("X2") + blue.ToString("X2");
        }

        private static int? GetTransformValue(OpenXmlElement transform) {
            string? value = transform.GetAttributes()
                .FirstOrDefault(attribute => string.Equals(attribute.LocalName, "val", StringComparison.Ordinal))
                .Value;
            return int.TryParse(value, out int result) ? result : null;
        }

        private static bool TryParseHexColor(string hexColor, out int red, out int green, out int blue) {
            red = 0;
            green = 0;
            blue = 0;
            if (hexColor.Length != 6) {
                return false;
            }

            try {
                red = Convert.ToInt32(hexColor.Substring(0, 2), 16);
                green = Convert.ToInt32(hexColor.Substring(2, 2), 16);
                blue = Convert.ToInt32(hexColor.Substring(4, 2), 16);
                return true;
            } catch (FormatException) {
                return false;
            }
        }

        private static int ClampColor(double value) {
            return (int)Math.Max(0D, Math.Min(255D, Math.Round(value)));
        }

        private static PowerPointSlideBackground GetBackgroundImage(A.BlipFill blipFill, OpenXmlPart ownerPart) {
            string? relationshipId = blipFill.Blip?.Embed?.Value;
            if (string.IsNullOrWhiteSpace(relationshipId)) {
                return PowerPointSlideBackground.Unsupported("The slide background image is missing its relationship id.");
            }

            ImagePart? imagePart;
            try {
                imagePart = ownerPart.GetPartById(relationshipId!) as ImagePart;
            } catch (ArgumentOutOfRangeException) {
                return PowerPointSlideBackground.Unsupported("The slide background image relationship could not be resolved.");
            }

            if (imagePart == null) {
                return PowerPointSlideBackground.Unsupported("The slide background relationship does not point to an image part.");
            }

            using Stream source = imagePart.GetStream(FileMode.Open, FileAccess.Read);
            using var copy = new MemoryStream();
            source.CopyTo(copy);
            PowerPointPictureCrop crop = ReadSourceCrop(blipFill.SourceRectangle);
            if (blipFill.GetFirstChild<A.Tile>() != null) {
                return PowerPointSlideBackground.Unsupported("The slide background image uses tiled placement, which is not currently supported by OfficeIMO exporters.");
            }

            return PowerPointSlideBackground.Image(copy.ToArray(), imagePart.ContentType, crop);
        }

        private static PowerPointPictureCrop ReadSourceCrop(A.SourceRectangle? rect) {
            if (rect == null) {
                return PowerPointPictureCrop.None;
            }

            return new PowerPointPictureCrop(
                ToCropFraction(rect.Left?.Value),
                ToCropFraction(rect.Top?.Value),
                ToCropFraction(rect.Right?.Value),
                ToCropFraction(rect.Bottom?.Value));
        }

        private static double ToCropFraction(int? value) {
            if (!value.HasValue) {
                return 0D;
            }

            return Math.Min(0.999999D, Math.Max(0D, value.Value / 100000D));
        }

        private static PowerPointSlideBackground GetBackgroundGradient(A.GradientFill gradientFill, A.ColorScheme? colorScheme, A.SchemeColor? placeholderColor) {
            A.GradientStop[] stops = gradientFill.GetFirstChild<A.GradientStopList>()?
                .Elements<A.GradientStop>()
                .OrderBy(stop => stop.Position?.Value ?? 0)
                .ToArray() ?? Array.Empty<A.GradientStop>();

            if (stops.Length < 2) {
                return PowerPointSlideBackground.Unsupported("The slide background gradient has fewer than two stops.");
            }

            string? startColor = ResolveGradientStopColor(stops[0], colorScheme, placeholderColor);
            string? endColor = ResolveGradientStopColor(stops[stops.Length - 1], colorScheme, placeholderColor);
            if (string.IsNullOrWhiteSpace(startColor) || string.IsNullOrWhiteSpace(endColor)) {
                return PowerPointSlideBackground.Unsupported("The slide background gradient uses colors that could not be resolved.");
            }

            double angleDegrees = (gradientFill.GetFirstChild<A.LinearGradientFill>()?.Angle?.Value ?? 0) / 60000D;
            return PowerPointSlideBackground.LinearGradient(startColor!, endColor!, angleDegrees);
        }

        private static string? ResolveGradientStopColor(A.GradientStop stop, A.ColorScheme? colorScheme, A.SchemeColor? placeholderColor) {
            A.RgbColorModelHex? rgbColor = stop.GetFirstChild<A.RgbColorModelHex>();
            string? rgbValue = rgbColor?.Val?.Value;
            if (!string.IsNullOrWhiteSpace(rgbValue)) {
                return ApplyColorTransforms(rgbValue, rgbColor);
            }

            A.SchemeColor? schemeColor = stop.GetFirstChild<A.SchemeColor>();
            string? scheme = GetSchemeColorValue(schemeColor);
            if (IsPlaceholderSchemeColor(scheme)) {
                string? placeholderScheme = GetSchemeColorValue(placeholderColor);
                string? placeholderResolvedColor = ResolveSchemeColor(colorScheme, placeholderScheme);
                placeholderResolvedColor = ApplyColorTransforms(placeholderResolvedColor, placeholderColor);
                return ApplyColorTransforms(placeholderResolvedColor, schemeColor);
            }

            return ApplyColorTransforms(ResolveSchemeColor(colorScheme, scheme), schemeColor);
        }
    }
}
