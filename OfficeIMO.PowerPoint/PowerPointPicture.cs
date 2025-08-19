using System;
using System.IO;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;
using ImagePartType = DocumentFormat.OpenXml.Packaging.PartTypeInfo;

namespace OfficeIMO.PowerPoint {
    /// <summary>
    /// Represents an image placed on a slide.
    /// </summary>
    public class PowerPointPicture : PowerPointShape {
        private readonly SlidePart _slidePart;

        internal PowerPointPicture(Picture picture, SlidePart slidePart) : base(picture) {
            _slidePart = slidePart;
        }

        /// <summary>
        /// Gets the MIME content type of the underlying image.
        /// </summary>
        public string? ContentType => GetImagePart()?.ContentType;

        /// <summary>
        /// Gets the MIME content type of the underlying image.
        /// </summary>
        public string? MimeType => ContentType;

        private ImagePart? GetImagePart() {
            Picture picture = (Picture)Element;
            string? relationshipId = picture.BlipFill?.Blip?.Embed?.Value;
            return relationshipId != null ? _slidePart.GetPartById(relationshipId) as ImagePart : null;
        }

        /// <summary>
        /// Replaces the picture's underlying image with the provided stream.
        /// </summary>
        /// <param name="newImage">Stream containing the new image data.</param>
        /// <param name="type">Image format of the new image.</param>
        public void UpdateImage(Stream newImage, ImagePartType type) {
            if (newImage == null) {
                throw new ArgumentNullException(nameof(newImage));
            }

            if (type != DocumentFormat.OpenXml.Packaging.ImagePartType.Png &&
                type != DocumentFormat.OpenXml.Packaging.ImagePartType.Jpeg &&
                type != DocumentFormat.OpenXml.Packaging.ImagePartType.Gif &&
                type != DocumentFormat.OpenXml.Packaging.ImagePartType.Bmp) {
                throw new NotSupportedException($"Image type {type} is not supported.");
            }

            Picture picture = (Picture)Element;
            A.Blip blip = picture.BlipFill?.Blip ?? throw new InvalidOperationException("Picture has no image");

            if (blip.Embed != null) {
                OpenXmlPart? oldPart = _slidePart.GetPartById(blip.Embed!);
                if (oldPart != null) {
                    _slidePart.DeletePart(oldPart);
                }
            }

            ImagePart imagePart = _slidePart.AddImagePart(type);
            newImage.Position = 0;
            imagePart.FeedData(newImage);
            string relId = _slidePart.GetIdOfPart(imagePart);
            blip.Embed = relId;
        }
    }
}

