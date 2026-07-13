using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using OfficeIMO.Drawing;
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
                OfficeColor? color = OfficeOpenXmlThemeColorResolver.ResolveColor(solidFill, colorScheme);
                solidColor = color.HasValue ? FormatBackgroundColor(color.Value) : null;
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
            OfficeColor? color = OfficeOpenXmlThemeColorResolver.ResolveColor(fill as A.SolidFill ?? fill.GetFirstChild<A.SolidFill>(), colorScheme, styleReference.GetFirstChild<A.SchemeColor>());
            string? solidColor = color.HasValue ? FormatBackgroundColor(color.Value) : null;
            if (!string.IsNullOrWhiteSpace(solidColor)) {
                return PowerPointSlideBackground.SolidColor(solidColor!);
            }

            A.GradientFill? gradientFill = fill as A.GradientFill ?? fill.GetFirstChild<A.GradientFill>();
            if (gradientFill != null) {
                return GetBackgroundGradient(gradientFill, colorScheme, styleReference.GetFirstChild<A.SchemeColor>());
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
            A.LinearGradientFill? linearGradientFill = gradientFill.GetFirstChild<A.LinearGradientFill>();
            if (linearGradientFill == null) {
                return PowerPointSlideBackground.Unsupported("The slide background uses a path or radial gradient, which is not currently supported by OfficeIMO exporters.");
            }

            A.GradientStop[] stops = gradientFill.GetFirstChild<A.GradientStopList>()?
                .Elements<A.GradientStop>()
                .OrderBy(stop => stop.Position?.Value ?? 0)
                .ToArray() ?? Array.Empty<A.GradientStop>();

            if (stops.Length < 2) {
                return PowerPointSlideBackground.Unsupported("The slide background gradient has fewer than two stops.");
            }

            OfficeColor? startColor = OfficeOpenXmlThemeColorResolver.ResolveColor(stops[0], colorScheme, placeholderColor);
            OfficeColor? endColor = OfficeOpenXmlThemeColorResolver.ResolveColor(stops[stops.Length - 1], colorScheme, placeholderColor);
            if (!startColor.HasValue || !endColor.HasValue) {
                return PowerPointSlideBackground.Unsupported("The slide background gradient uses colors that could not be resolved.");
            }

            double angleDegrees = (linearGradientFill.Angle?.Value ?? 0) / 60000D;
            return PowerPointSlideBackground.LinearGradient(FormatBackgroundColor(startColor.Value), FormatBackgroundColor(endColor.Value), angleDegrees);
        }

        private static string FormatBackgroundColor(OfficeColor color) =>
            color.A == 255 ? color.ToRgbHex() : color.ToHex();

    }
}
