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
            return GetBackgroundCore(maximumImageBytes: null);
        }

        /// <summary>Returns a detached slide-background snapshot with a bound on embedded image bytes.</summary>
        public PowerPointSlideBackground GetBackground(int maximumImageBytes) {
            if (maximumImageBytes <= 0) throw new ArgumentOutOfRangeException(nameof(maximumImageBytes));
            return GetBackgroundCore(maximumImageBytes);
        }

        /// <summary>Evaluates an aggregate image allowance only when the background contains an image.</summary>
        internal PowerPointSlideBackground GetBackground(Func<int> imageReadLimit) {
            if (imageReadLimit == null) throw new ArgumentNullException(nameof(imageReadLimit));
            return GetBackgroundCore(maximumImageBytes: null, imageReadLimit);
        }

        /// <summary>Resolves fill colors for semantic exporters without opening or copying image part streams.</summary>
        internal PowerPointSlideBackground GetBackgroundColors() => GetBackgroundCore(maximumImageBytes: null, includeImageData: false);

        private PowerPointSlideBackground GetBackgroundCore(int? maximumImageBytes, Func<int>? imageReadLimit = null, bool includeImageData = true) {
            (BackgroundProperties? properties, OpenXmlPart? ownerPart) = GetResolvedBackgroundProperties();
            if (properties == null || !properties.HasChildren) {
                BackgroundStyleReference? styleReference = GetResolvedBackgroundStyleReference().StyleReference;
                return styleReference != null
                    ? GetBackgroundStyleReference(styleReference)
                    : PowerPointSlideBackground.None();
            }

            if (properties.GetFirstChild<A.NoFill>() != null) return PowerPointSlideBackground.None();

            A.SolidFill? solidFill = properties.GetFirstChild<A.SolidFill>();
            string? solidColor = null;
            if (solidFill != null) {
                A.ColorScheme? colorScheme = ResolvedColorScheme;
                OfficeColor? color = OfficeOpenXmlThemeColorResolver.ResolveColor(solidFill, colorScheme, colorMap: ResolvedColorMap);
                solidColor = color.HasValue ? FormatBackgroundColor(color.Value) : null;
            }

            if (!string.IsNullOrWhiteSpace(solidColor)) {
                return PowerPointSlideBackground.SolidColor(solidColor!);
            }

            A.BlipFill? blipFill = properties.GetFirstChild<A.BlipFill>();
            if (blipFill != null) {
                return includeImageData
                    ? GetBackgroundImage(blipFill, ownerPart ?? _slidePart, imageReadLimit?.Invoke() ?? maximumImageBytes)
                    : PowerPointSlideBackground.None();
            }

            A.GradientFill? gradientFill = properties.GetFirstChild<A.GradientFill>();
            if (gradientFill != null) {
                A.ColorScheme? colorScheme = ResolvedColorScheme;
                return GetBackgroundGradient(gradientFill, colorScheme, placeholderColor: null, colorMap: ResolvedColorMap);
            }

            return PowerPointSlideBackground.Unsupported("The slide background fill type is not currently supported by OfficeIMO exporters.");
        }

        private (BackgroundProperties? Properties, OpenXmlPart? OwnerPart) GetResolvedBackgroundProperties() {
            Background? slideBackground = SlideRoot.CommonSlideData?.Background;
            BackgroundProperties? slideProperties = slideBackground?.BackgroundProperties;
            if (slideBackground?.BackgroundStyleReference != null || slideProperties?.HasChildren == true) {
                return (slideProperties, _slidePart);
            }

            SlideLayoutPart? layoutPart = _slidePart.SlideLayoutPart;
            Background? layoutBackground = layoutPart?.SlideLayout?.CommonSlideData?.Background;
            BackgroundProperties? layoutProperties = layoutBackground?.BackgroundProperties;
            if (layoutBackground?.BackgroundStyleReference != null || layoutProperties?.HasChildren == true) {
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

        private PowerPointSlideBackground GetBackgroundStyleReference(BackgroundStyleReference styleReference) {
            A.FormatScheme? formatScheme = ResolvedFormatScheme;
            if (formatScheme == null) {
                return PowerPointSlideBackground.Unsupported("The slide background references a theme background style but no theme format scheme could be resolved.");
            }

            OpenXmlElement? fill = ResolveThemeBackgroundFill(formatScheme, styleReference.Index?.Value);
            if (fill == null) {
                return PowerPointSlideBackground.Unsupported("The slide background references a theme background style that could not be resolved.");
            }

            A.ColorScheme? colorScheme = ResolvedColorScheme;
            OfficeColor? color = OfficeOpenXmlThemeColorResolver.ResolveColor(
                fill as A.SolidFill ?? fill.GetFirstChild<A.SolidFill>(),
                colorScheme, styleReference, ResolvedColorMap);
            string? solidColor = color.HasValue ? FormatBackgroundColor(color.Value) : null;
            if (!string.IsNullOrWhiteSpace(solidColor)) {
                return PowerPointSlideBackground.SolidColor(solidColor!);
            }

            A.GradientFill? gradientFill = fill as A.GradientFill ?? fill.GetFirstChild<A.GradientFill>();
            if (gradientFill != null) {
                return GetBackgroundGradient(gradientFill, colorScheme,
                    styleReference, ResolvedColorMap);
            }

            return PowerPointSlideBackground.Unsupported("The slide background references a theme background fill type that is not currently supported by OfficeIMO exporters.");
        }

        private static OpenXmlElement? ResolveThemeBackgroundFill(A.FormatScheme formatScheme, uint? index) {
            if (!index.HasValue || index.Value == 0U || index.Value == 1000U) return null;
            bool background = index.Value >= 1001U;
            OpenXmlElementList fills = (background
                ? formatScheme.GetFirstChild<A.BackgroundFillStyleList>()?.ChildElements
                : formatScheme.GetFirstChild<A.FillStyleList>()?.ChildElements) ?? default;
            uint zeroBased = index.Value - (background ? 1001U : 1U);
            return zeroBased < unchecked((uint)fills.Count) ? fills[(int)zeroBased] : null;
        }

        private static PowerPointSlideBackground GetBackgroundImage(
            A.BlipFill blipFill,
            OpenXmlPart ownerPart,
            int? maximumImageBytes) {
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
            byte[] imageBytes = maximumImageBytes.HasValue
                ? OfficeIMO.Core.Internal.OfficeStreamReader.ReadAllBytes(source, maximumImageBytes.Value)
                : OfficeIMO.Core.Internal.OfficeStreamReader.ReadAllBytes(source);
            PowerPointPictureCrop crop = ReadSourceCrop(blipFill.SourceRectangle);
            if (blipFill.GetFirstChild<A.Tile>() != null) {
                return PowerPointSlideBackground.Unsupported("The slide background image uses tiled placement, which is not currently supported by OfficeIMO exporters.");
            }

            return PowerPointSlideBackground.Image(imageBytes, imagePart.ContentType, crop);
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

        private static PowerPointSlideBackground GetBackgroundGradient(
            A.GradientFill gradientFill, A.ColorScheme? colorScheme,
            OpenXmlElement? placeholderColor, OpenXmlElement? colorMap) {
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

            OfficeColor? startColor = OfficeOpenXmlThemeColorResolver.ResolveColor(stops[0], colorScheme, placeholderColor, colorMap);
            OfficeColor? endColor = OfficeOpenXmlThemeColorResolver.ResolveColor(stops[stops.Length - 1], colorScheme, placeholderColor, colorMap);
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
