using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.Drawing;
using OfficeIMO.OpenXml.Internal;
using OfficeIMO.PowerPoint.OpenXml;

namespace OfficeIMO.PowerPoint {
    public partial class PowerPointSlide {
        private const int BackgroundPml = OfficeXmlNamespaces.PresentationML;
        private const int BackgroundDrawing = OfficeXmlNamespaces.DrawingML;

        /// <summary>
        ///     Gets or sets the local solid background color, retaining native theme slots and opacity.
        /// </summary>
        public OfficeColor? BackgroundColor {
            get => PowerPointDrawingColor.Read(BackgroundOf(PartRef).Child(BackgroundPml, "bgPr").Child(BackgroundDrawing, "solidFill"));
            set {
                if (value is not { } color) {
                    PowerPointPartRef current = PartRef;
                    RemoveBackgroundFill(current, EnsureCommonSlideData(current), "solidFill");
                    return;
                }

                OfficeTheme? theme = PowerPointDrawingColor.NeedsTheme(color) ? AuthoringTheme : null;
                ClearBackgroundImage();
                PowerPointPartRef part = PartRef;
                OfficeXmlNode data = EnsureCommonSlideData(part);
                SetBackgroundFill(part, data, PowerPointDrawingColor.CreateFill(data, color, theme));
            }
        }

        /// <summary>Sets an explicit no-fill slide background, overriding inherited theme fills.</summary>
        public void SetBackgroundNoFill() {
            PowerPointPartRef part = PartRef;
            OfficeXmlNode data = EnsureCommonSlideData(part);
            OfficeXmlNode background = data.CreateElement(BackgroundPml, "bg");
            background.Append(background.CreateElement(BackgroundPml, "bgPr"))
                .Append(background.CreateElement(BackgroundDrawing, "noFill"));
            data.RemoveChildren(BackgroundPml, "bg");
            data.InsertOrdered(background, PCommonSlideDataView.Order);
            part.MarkEdited();
        }

        /// <summary>
        ///     Sets a background image for the slide.
        /// </summary>
        public void SetBackgroundImage(string imagePath) {
            ArgumentNullException.ThrowIfNull(imagePath);
            if (!File.Exists(imagePath)) {
                throw new FileNotFoundException("Image file not found.", imagePath);
            }

            OfficeIMO.Drawing.OfficeImageFormat imageType = GetImagePartType(imagePath);
            using FileStream stream = new(imagePath, FileMode.Open, FileAccess.Read);
            using MemoryStream? normalized = PowerPointImageNormalization.NormalizeHeicForEmbedding(stream, imageType);
            if (normalized != null) {
                SetBackgroundImage(normalized, OfficeIMO.Drawing.OfficeImageFormat.Png);
                return;
            }
            string imageExtension = PowerPointPartFactory.GetImageExtension(imageType, imagePath);
            SetBackgroundImageCore(stream, imageType, imageExtension);
        }

        /// <summary>
        /// Sets a background image for the slide from a readable image stream.
        /// </summary>
        /// <param name="image">Stream containing the image data.</param>
        /// <param name="imageType">Image format of the stream.</param>
        public void SetBackgroundImage(Stream image, OfficeIMO.Drawing.OfficeImageFormat imageType) {
            ArgumentNullException.ThrowIfNull(image);
            if (!image.CanRead) {
                throw new ArgumentException("Image stream must be readable.", nameof(image));
            }

            using MemoryStream? normalized = PowerPointImageNormalization.NormalizeHeicForEmbedding(image, imageType);
            if (normalized != null) {
                SetBackgroundImage(normalized, OfficeIMO.Drawing.OfficeImageFormat.Png);
                return;
            }
            SetBackgroundImageCore(image, imageType, PowerPointPartFactory.GetImageExtension(imageType));
        }

        private void SetBackgroundImageCore(Stream image, OfficeImageFormat imageType, string imageExtension) {
            string relationshipId;
            PowerPointPartRef part = PartRef;
            if (!part.IsSdk && OwnerDocument is { CanEditPackage: true } document) {
                if (image.CanSeek) image.Position = 0;
                using var buffer = new MemoryStream();
                image.CopyTo(buffer);
                if (!document.TryAddPackageImagePart(part.Name, buffer.ToArray(), PowerPointImageContentTypes.For(imageType),
                        imageExtension, out relationshipId, out _, out _, reuse: false)) {
                    throw new InvalidOperationException("The slide image part could not be added.");
                }
            } else {
                relationshipId = AddSdkBackgroundImagePart(image, imageType, imageExtension);
                part = PartRef;
            }

            OfficeXmlNode data = EnsureCommonSlideData(part);
            string? previousRelationshipId = BackgroundImageRelationship(data);
            OfficeXmlNode fill = data.CreateElement(BackgroundDrawing, "blipFill");
            fill.Append(fill.CreateElement(BackgroundDrawing, "blip"))
                .SetAttribute(OfficeXmlNamespaces.Relationships, "embed", relationshipId);
            fill.Append(fill.CreateElement(BackgroundDrawing, "stretch")).Append(fill.CreateElement(BackgroundDrawing, "fillRect"));
            SetBackgroundFill(part, data, fill);
            if (previousRelationshipId != null) part.RemoveRelationshipIfUnused(previousRelationshipId, default);
        }

        /// <summary>
        ///     Sets a linear gradient background, retaining theme roles and color opacity.
        /// </summary>
        public void SetBackgroundGradient(OfficeColor startColor, OfficeColor endColor, double angleDegrees = 135d) {
            int angle = ToOpenXmlAngle(angleDegrees);
            OfficeTheme? theme = PowerPointDrawingColor.NeedsTheme(startColor) || PowerPointDrawingColor.NeedsTheme(endColor)
                ? AuthoringTheme : null;
            ClearBackgroundImage();
            PowerPointPartRef part = PartRef;
            OfficeXmlNode data = EnsureCommonSlideData(part);
            OfficeXmlNode gradient = data.CreateElement(BackgroundDrawing, "gradFill");
            gradient.SetBooleanAttribute("rotWithShape", true);
            OfficeXmlNode stops = gradient.Append(gradient.CreateElement(BackgroundDrawing, "gsLst"));
            OfficeXmlNode first = stops.Append(stops.CreateElement(BackgroundDrawing, "gs"));
            first.SetInt32Attribute("pos", 0);
            PowerPointDrawingColor.Append(first, startColor, theme);
            OfficeXmlNode last = stops.Append(stops.CreateElement(BackgroundDrawing, "gs"));
            last.SetInt32Attribute("pos", 100000);
            PowerPointDrawingColor.Append(last, endColor, theme);
            OfficeXmlNode linear = gradient.Append(gradient.CreateElement(BackgroundDrawing, "lin"));
            linear.SetInt32Attribute("ang", angle);
            linear.SetBooleanAttribute("scaled", false);
            SetBackgroundFill(part, data, gradient);
        }

        /// <summary>
        ///     Clears any background image from the slide.
        /// </summary>
        public void ClearBackgroundImage() {
            PowerPointPartRef part = PartRef;
            OfficeXmlNode data = part.Root.Child(BackgroundPml, "cSld");
            string? previousRelationshipId = BackgroundImageRelationship(data);
            RemoveBackgroundFill(part, data, "blipFill");
            if (previousRelationshipId != null) part.RemoveRelationshipIfUnused(previousRelationshipId, default);
        }

        private static string? BackgroundImageRelationship(OfficeXmlNode data) =>
            data.Child(BackgroundPml, "bg").Child(BackgroundPml, "bgPr").Child(BackgroundDrawing, "blipFill")
                .Child(BackgroundDrawing, "blip").Attribute(OfficeXmlNamespaces.Relationships, "embed");

        private static void RemoveBackgroundFill(PowerPointPartRef part, OfficeXmlNode data, string name) {
            OfficeXmlNode background = data.Child(BackgroundPml, "bg");
            OfficeXmlNode properties = background.Child(BackgroundPml, "bgPr");
            if (properties.IsNone) return;
            properties.RemoveChildren(BackgroundDrawing, name);
            if (properties.FirstChild.IsNone) background.Remove();
            part.MarkEdited();
        }

        private static void SetBackgroundFill(PowerPointPartRef part, OfficeXmlNode data, OfficeXmlNode fill) {
            OfficeXmlNode background = data.GetOrAddChild(BackgroundPml, "bg", PCommonSlideDataView.Order);
            OfficeXmlNode properties = background.GetOrPrependChild(BackgroundPml, "bgPr");
            properties.RemoveChildren(BackgroundDrawing, "blipFill");
            properties.RemoveChildren(BackgroundDrawing, "gradFill");
            properties.RemoveChildren(BackgroundDrawing, "grpFill");
            properties.RemoveChildren(BackgroundDrawing, "noFill");
            properties.RemoveChildren(BackgroundDrawing, "pattFill");
            properties.RemoveChildren(BackgroundDrawing, "solidFill");
            properties.Append(fill);
            part.MarkEdited();
        }

        private static int ToOpenXmlAngle(double degrees) {
            double normalized = degrees % 360d;
            if (normalized < 0) normalized += 360d;
            return (int)Math.Round(normalized * 60000d);
        }
    }
}
