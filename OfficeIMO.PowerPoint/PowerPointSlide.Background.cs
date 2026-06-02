using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using P14 = DocumentFormat.OpenXml.Office2010.PowerPoint;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.PowerPoint {
    public partial class PowerPointSlide {
        /// <summary>
        ///     Gets or sets the slide background color in hex format (e.g. "FF0000").
        /// </summary>
        public string? BackgroundColor {
            get {
                CommonSlideData? common = SlideRoot.CommonSlideData;
                Background? bg = common?.Background;
                A.SolidFill? solid = bg?.BackgroundProperties?.GetFirstChild<A.SolidFill>();
                return solid?.RgbColorModelHex?.Val;
            }
            set {
                CommonSlideData common = SlideRoot.CommonSlideData ??= new CommonSlideData(new ShapeTree());
                if (value == null) {
                    BackgroundProperties? properties = common.Background?.BackgroundProperties;
                    if (properties == null) {
                        return;
                    }

                    properties.RemoveAllChildren<A.SolidFill>();
                    if (!properties.HasChildren) {
                        common.Background = null;
                    }
                    return;
                }

                Background bg = common.Background ?? new Background();
                BackgroundProperties props = bg.BackgroundProperties ?? new BackgroundProperties();
                RemoveBackgroundFillChildren(props);
                props.Append(new A.SolidFill(new A.RgbColorModelHex { Val = value }));
                bg.BackgroundProperties = props;
                common.Background = bg;
            }
        }

        /// <summary>
        ///     Sets a background image for the slide.
        /// </summary>
        public void SetBackgroundImage(string imagePath) {
            if (imagePath == null) {
                throw new ArgumentNullException(nameof(imagePath));
            }
            if (!File.Exists(imagePath)) {
                throw new FileNotFoundException("Image file not found.", imagePath);
            }

            A.Blip? previousBlip = GetBackgroundBlip();
            string? previousRelationshipId = previousBlip?.Embed?.Value;
            ImagePartType imageType = GetImagePartType(imagePath);
            PartTypeInfo partTypeInfo = imageType.ToPartTypeInfo();
            string imageExtension = PowerPointPartFactory.GetImageExtension(imageType, imagePath);
            string imagePartUri = PowerPointPartFactory.GetIndexedPartUri(
                _slidePart.OpenXmlPackage,
                "ppt/media",
                "image",
                imageExtension,
                allowBaseWithoutIndex: false);

            ImagePart imagePart = PowerPointPartFactory.CreatePart<ImagePart>(
                _slidePart,
                partTypeInfo.ContentType,
                imagePartUri);

            using FileStream stream = new(imagePath, FileMode.Open, FileAccess.Read);
            imagePart.FeedData(stream);
            string relationshipId = _slidePart.GetIdOfPart(imagePart);

            CommonSlideData common = SlideRoot.CommonSlideData ??= new CommonSlideData(new ShapeTree());
            Background background = common.Background ?? new Background();
            BackgroundProperties props = background.BackgroundProperties ?? new BackgroundProperties();

            RemoveBackgroundFillChildren(props);

            props.Append(new A.BlipFill(
                new A.Blip { Embed = relationshipId },
                new A.Stretch(new A.FillRectangle())
            ));

            background.BackgroundProperties = props;
            common.Background = background;

            RemoveUnusedImagePart(previousRelationshipId, previousBlip);
        }

        /// <summary>
        ///     Sets a linear gradient background for the slide using two 6-digit hex colors.
        /// </summary>
        public void SetBackgroundGradient(string startColor, string endColor, double angleDegrees = 135d) {
            if (string.IsNullOrWhiteSpace(startColor)) {
                throw new ArgumentException("Gradient start color cannot be null or empty.", nameof(startColor));
            }

            if (string.IsNullOrWhiteSpace(endColor)) {
                throw new ArgumentException("Gradient end color cannot be null or empty.", nameof(endColor));
            }

            string normalizedStart = NormalizeHexColor(startColor);
            string normalizedEnd = NormalizeHexColor(endColor);

            CommonSlideData common = SlideRoot.CommonSlideData ??= new CommonSlideData(new ShapeTree());
            Background background = common.Background ?? new Background();
            BackgroundProperties props = background.BackgroundProperties ?? new BackgroundProperties();

            RemoveBackgroundFillChildren(props);

            A.GradientFill gradient = new() { RotateWithShape = true };
            A.GradientStopList stops = new();
            stops.Append(
                new A.GradientStop(new A.RgbColorModelHex { Val = normalizedStart }) {
                    Position = 0
                });
            stops.Append(
                new A.GradientStop(new A.RgbColorModelHex { Val = normalizedEnd }) {
                    Position = 100000
                });
            gradient.Append(stops);
            gradient.Append(new A.LinearGradientFill {
                Angle = ToOpenXmlAngle(angleDegrees),
                Scaled = false
            });

            props.Append(gradient);
            background.BackgroundProperties = props;
            common.Background = background;
        }

        /// <summary>
        ///     Clears any background image from the slide.
        /// </summary>
        public void ClearBackgroundImage() {
            CommonSlideData? common = SlideRoot.CommonSlideData;
            if (common?.Background?.BackgroundProperties == null) {
                return;
            }

            A.Blip? previousBlip = GetBackgroundBlip();
            string? previousRelationshipId = previousBlip?.Embed?.Value;
            common.Background.BackgroundProperties.RemoveAllChildren<A.BlipFill>();
            if (!common.Background.BackgroundProperties.HasChildren) {
                common.Background = null;
            }

            RemoveUnusedImagePart(previousRelationshipId, previousBlip);
        }

        private static void RemoveBackgroundFillChildren(BackgroundProperties properties) {
            properties.RemoveAllChildren<A.BlipFill>();
            properties.RemoveAllChildren<A.GradientFill>();
            properties.RemoveAllChildren<A.GroupFill>();
            properties.RemoveAllChildren<A.NoFill>();
            properties.RemoveAllChildren<A.PatternFill>();
            properties.RemoveAllChildren<A.SolidFill>();
        }

        private static string NormalizeHexColor(string value) {
            string normalized = value.Trim();
            if (normalized.StartsWith("#", StringComparison.Ordinal)) {
                normalized = normalized.Substring(1);
            }

            if (normalized.Length != 6 || normalized.Any(c => !Uri.IsHexDigit(c))) {
                throw new ArgumentException("Color must be a 6-digit hex value.", nameof(value));
            }

            return normalized.ToUpperInvariant();
        }

        private static int ToOpenXmlAngle(double degrees) {
            double normalized = degrees % 360d;
            if (normalized < 0) {
                normalized += 360d;
            }

            return (int)Math.Round(normalized * 60000d);
        }

        private A.Blip? GetBackgroundBlip() {
            return SlideRoot.CommonSlideData?.Background?.BackgroundProperties?.GetFirstChild<A.BlipFill>()?.Blip;
        }

        private void RemoveUnusedImagePart(string? relationshipId, A.Blip? currentBlip) {
            string resolvedRelationshipId = relationshipId ?? string.Empty;
            if (string.IsNullOrWhiteSpace(resolvedRelationshipId)) {
                return;
            }
            if (IsImageRelationshipReferenced(resolvedRelationshipId, currentBlip)) {
                return;
            }

            try {
                _slidePart.DeletePart(resolvedRelationshipId);
            } catch (ArgumentOutOfRangeException) {
                // The previous relationship may already be absent on damaged input.
            }
        }

        private bool IsImageRelationshipReferenced(string relationshipId, A.Blip? currentBlip) {
            return SlideRoot
                .Descendants<A.Blip>()
                .Any(blip => !ReferenceEquals(blip, currentBlip) && blip.Embed?.Value == relationshipId);
        }
    }
}
