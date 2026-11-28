using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.Drawing;
using OfficeIMO.OpenXml.Internal;
using OfficeIMO.PowerPoint.OpenXml;

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
            ArgumentOutOfRangeException.ThrowIfNegativeOrZero(maximumImageBytes);
            return GetBackgroundCore(maximumImageBytes);
        }

        /// <summary>Evaluates an aggregate image allowance only when the background contains an image.</summary>
        internal PowerPointSlideBackground GetBackground(Func<int> imageReadLimit) {
            ArgumentNullException.ThrowIfNull(imageReadLimit);
            return GetBackgroundCore(maximumImageBytes: null, imageReadLimit);
        }

        /// <summary>Resolves fill colors for semantic exporters without opening or copying image part streams.</summary>
        internal PowerPointSlideBackground GetBackgroundColors() => GetBackgroundCore(maximumImageBytes: null, includeImageData: false);

        private PowerPointSlideBackground GetBackgroundCore(int? maximumImageBytes, Func<int>? imageReadLimit = null, bool includeImageData = true) {
            PowerPointPartRef slide = PartRef;
            PowerPointPartRef layout = slide.RelatedPart(PowerPointPartRef.SlideLayoutRelationshipType);
            PowerPointPartRef master = layout.RelatedPart(PowerPointPartRef.SlideMasterRelationshipType);
            OfficeXmlNode slideBackground = BackgroundOf(slide);
            OfficeXmlNode layoutBackground = BackgroundOf(layout);
            OfficeXmlNode masterBackground = BackgroundOf(master);
            OfficeXmlNode properties;
            PowerPointPartRef owner;
            if (DefinesBackground(slideBackground)) {
                properties = slideBackground.Child(OfficeXmlNamespaces.PresentationML, "bgPr");
                owner = slide;
            } else if (DefinesBackground(layoutBackground)) {
                properties = layoutBackground.Child(OfficeXmlNamespaces.PresentationML, "bgPr");
                owner = layout;
            } else {
                properties = masterBackground.Child(OfficeXmlNamespaces.PresentationML, "bgPr");
                owner = master;
            }

            if (properties.FirstChild.IsNone) {
                OfficeXmlNode reference = slideBackground.Child(OfficeXmlNamespaces.PresentationML, "bgRef");
                if (reference.IsNone) reference = layoutBackground.Child(OfficeXmlNamespaces.PresentationML, "bgRef");
                if (reference.IsNone) reference = masterBackground.Child(OfficeXmlNamespaces.PresentationML, "bgRef");
                return reference.IsNone
                    ? PowerPointSlideBackground.None()
                    : GetBackgroundStyleReference(reference, PowerPointThemeScope.For(slide));
            }

            const int drawing = OfficeXmlNamespaces.DrawingML;
            if (properties.Child(drawing, "noFill").Exists) return PowerPointSlideBackground.None();

            OfficeXmlNode solid = properties.Child(drawing, "solidFill");
            if (solid.Exists) {
                PowerPointThemeScope scope = PowerPointThemeScope.For(slide);
                OfficeColor? color = OfficeOpenXmlThemeColorResolver.ResolveColor(solid, scope.ColorScheme, colorMap: scope.ColorMap);
                if (color.HasValue) return PowerPointSlideBackground.SolidColor(FormatBackgroundColor(color.Value));
            }

            OfficeXmlNode picture = properties.Child(drawing, "blipFill");
            if (picture.Exists) {
                return includeImageData
                    ? GetBackgroundImage(picture, owner, imageReadLimit?.Invoke() ?? maximumImageBytes)
                    : PowerPointSlideBackground.None();
            }

            OfficeXmlNode gradient = properties.Child(drawing, "gradFill");
            return gradient.Exists
                ? GetBackgroundGradient(gradient, PowerPointThemeScope.For(slide))
                : PowerPointSlideBackground.Unsupported("The slide background fill type is not currently supported by OfficeIMO exporters.");
        }

        private static OfficeXmlNode BackgroundOf(PowerPointPartRef part) =>
            part.Root.Child(OfficeXmlNamespaces.PresentationML, "cSld").Child(OfficeXmlNamespaces.PresentationML, "bg");

        private static bool DefinesBackground(OfficeXmlNode background) =>
            background.Child(OfficeXmlNamespaces.PresentationML, "bgRef").Exists
            || background.Child(OfficeXmlNamespaces.PresentationML, "bgPr").FirstChild.Exists;

        private static PowerPointSlideBackground GetBackgroundStyleReference(OfficeXmlNode reference, PowerPointThemeScope scope) {
            if (scope.FormatScheme.IsNone) {
                return PowerPointSlideBackground.Unsupported("The slide background references a theme background style but no theme format scheme could be resolved.");
            }

            OfficeXmlNode fill = ResolveThemeBackgroundFill(scope.FormatScheme, reference.UInt32Attribute("idx"));
            if (fill.IsNone) {
                return PowerPointSlideBackground.Unsupported("The slide background references a theme background style that could not be resolved.");
            }

            const int drawing = OfficeXmlNamespaces.DrawingML;
            OfficeXmlNode solid = fill.Is(drawing, "solidFill") ? fill : fill.Child(drawing, "solidFill");
            OfficeColor? color = OfficeOpenXmlThemeColorResolver.ResolveColor(solid, scope.ColorScheme, reference, scope.ColorMap);
            if (color.HasValue) return PowerPointSlideBackground.SolidColor(FormatBackgroundColor(color.Value));

            OfficeXmlNode gradient = fill.Is(drawing, "gradFill") ? fill : fill.Child(drawing, "gradFill");
            return gradient.Exists
                ? GetBackgroundGradient(gradient, scope, reference)
                : PowerPointSlideBackground.Unsupported("The slide background references a theme background fill type that is not currently supported by OfficeIMO exporters.");
        }

        private static OfficeXmlNode ResolveThemeBackgroundFill(OfficeXmlNode formatScheme, uint? index) {
            if (!index.HasValue || index.Value == 0U || index.Value == 1000U) return default;
            bool background = index.Value >= 1001U;
            uint remaining = index.Value - (background ? 1001U : 1U);
            foreach (OfficeXmlNode fill in formatScheme.Child(OfficeXmlNamespaces.DrawingML,
                background ? "bgFillStyleLst" : "fillStyleLst").Children()) {
                if (remaining-- == 0U) return fill;
            }

            return default;
        }

        private static PowerPointSlideBackground GetBackgroundImage(
            OfficeXmlNode blipFill, PowerPointPartRef owner, int? maximumImageBytes) {
            const int drawing = OfficeXmlNamespaces.DrawingML;
            string? relationshipId = blipFill.Child(drawing, "blip").Attribute(OfficeXmlNamespaces.Relationships, "embed");
            if (string.IsNullOrWhiteSpace(relationshipId)) {
                return PowerPointSlideBackground.Unsupported("The slide background image is missing its relationship id.");
            }

            if (!owner.TryGetRelationship(relationshipId!, out PowerPointRelationship relationship) || relationship.TargetPart.IsNone) {
                return PowerPointSlideBackground.Unsupported("The slide background image relationship could not be resolved.");
            }

            PowerPointPartRef image = relationship.TargetPart;
            string? contentType = image.ContentType;
            if (OfficeXmlNamespaces.NormalizeRelationshipType(relationship.Type)
                != "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" || contentType == null) {
                return PowerPointSlideBackground.Unsupported("The slide background relationship does not point to an image part.");
            }

            byte[] imageBytes = image.ReadBytes(maximumImageBytes);
            OfficeXmlNode crop = blipFill.Child(drawing, "srcRect");
            if (blipFill.Child(drawing, "tile").Exists) {
                return PowerPointSlideBackground.Unsupported("The slide background image uses tiled placement, which is not currently supported by OfficeIMO exporters.");
            }

            return PowerPointSlideBackground.Image(imageBytes, contentType, new PowerPointImageCrop(
                ToCropFraction(crop.Int32Attribute("l")), ToCropFraction(crop.Int32Attribute("t")),
                ToCropFraction(crop.Int32Attribute("r")), ToCropFraction(crop.Int32Attribute("b"))));
        }

        private static double ToCropFraction(int? value) =>
            Math.Min(0.999999D, Math.Max(0D, (value ?? 0) / 100000D));

        private static PowerPointSlideBackground GetBackgroundGradient(
            OfficeXmlNode gradient, PowerPointThemeScope scope, OfficeXmlNode placeholderColor = default) {
            const int drawing = OfficeXmlNamespaces.DrawingML;
            OfficeXmlNode linear = gradient.Child(drawing, "lin");
            if (linear.IsNone) {
                return PowerPointSlideBackground.Unsupported("The slide background uses a path or radial gradient, which is not currently supported by OfficeIMO exporters.");
            }

            OfficeXmlNode first = default, last = default;
            int firstPosition = int.MaxValue, lastPosition = int.MinValue;
            foreach (OfficeXmlNode stop in gradient.Child(drawing, "gsLst").Children(drawing, "gs")) {
                int position = stop.Int32Attribute("pos") ?? 0;
                if (first.IsNone || position < firstPosition) {
                    first = stop;
                    firstPosition = position;
                }
                if (last.IsNone || position >= lastPosition) {
                    last = stop;
                    lastPosition = position;
                }
            }

            if (first == last) {
                return PowerPointSlideBackground.Unsupported("The slide background gradient has fewer than two stops.");
            }

            OfficeColor? startColor = OfficeOpenXmlThemeColorResolver.ResolveColor(first, scope.ColorScheme, placeholderColor, scope.ColorMap);
            OfficeColor? endColor = OfficeOpenXmlThemeColorResolver.ResolveColor(last, scope.ColorScheme, placeholderColor, scope.ColorMap);
            if (!startColor.HasValue || !endColor.HasValue) {
                return PowerPointSlideBackground.Unsupported("The slide background gradient uses colors that could not be resolved.");
            }

            double angle = (linear.Int32Attribute("ang") ?? 0) / 60000D;
            return PowerPointSlideBackground.LinearGradient(FormatBackgroundColor(startColor.Value), FormatBackgroundColor(endColor.Value), angle);
        }

        private static string FormatBackgroundColor(OfficeColor color) =>
            color.A == 255 ? color.ToRgbHex() : color.ToHex();

    }
}
