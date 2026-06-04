using System;
using System.IO;
using System.Linq;
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
                return PowerPointSlideBackground.None();
            }

            A.SolidFill? solidFill = properties.GetFirstChild<A.SolidFill>();
            string? solidColor = solidFill?.RgbColorModelHex?.Val?.Value;
            if (!string.IsNullOrWhiteSpace(solidColor)) {
                return PowerPointSlideBackground.SolidColor(solidColor!);
            }

            A.BlipFill? blipFill = properties.GetFirstChild<A.BlipFill>();
            if (blipFill != null) {
                return GetBackgroundImage(blipFill, ownerPart ?? _slidePart);
            }

            A.GradientFill? gradientFill = properties.GetFirstChild<A.GradientFill>();
            if (gradientFill != null) {
                return GetBackgroundGradient(gradientFill);
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

        private static PowerPointSlideBackground GetBackgroundGradient(A.GradientFill gradientFill) {
            A.GradientStop[] stops = gradientFill.GetFirstChild<A.GradientStopList>()?
                .Elements<A.GradientStop>()
                .OrderBy(stop => stop.Position?.Value ?? 0)
                .ToArray() ?? Array.Empty<A.GradientStop>();

            if (stops.Length < 2) {
                return PowerPointSlideBackground.Unsupported("The slide background gradient has fewer than two stops.");
            }

            string? startColor = stops[0].GetFirstChild<A.RgbColorModelHex>()?.Val?.Value;
            string? endColor = stops[stops.Length - 1].GetFirstChild<A.RgbColorModelHex>()?.Val?.Value;
            if (string.IsNullOrWhiteSpace(startColor) || string.IsNullOrWhiteSpace(endColor)) {
                return PowerPointSlideBackground.Unsupported("The slide background gradient uses non-RGB theme or scheme colors that are not yet resolved.");
            }

            double angleDegrees = (gradientFill.GetFirstChild<A.LinearGradientFill>()?.Angle?.Value ?? 0) / 60000D;
            return PowerPointSlideBackground.LinearGradient(startColor!, endColor!, angleDegrees);
        }
    }
}
