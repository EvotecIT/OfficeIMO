using System;
using System.IO;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.PowerPoint {
    /// <summary>
    ///     Represents an image placed on a slide.
    /// </summary>
    public class PowerPointPicture : PowerPointShape {
        private readonly SlidePart _slidePart;

        internal PowerPointPicture(Picture picture, SlidePart slidePart) : base(picture) {
            _slidePart = slidePart;
        }

        private Picture Picture => (Picture)Element;

        /// <summary>
        ///     Gets the MIME content type of the underlying image.
        /// </summary>
        public string? ContentType => GetImagePart()?.ContentType;

        /// <summary>
        ///     Gets the MIME content type of the underlying image.
        /// </summary>
        public string? MimeType => ContentType;

        private ImagePart? GetImagePart() {
            Picture picture = (Picture)Element;
            string? relationshipId = picture.BlipFill?.Blip?.Embed?.Value;
            return relationshipId != null ? _slidePart.GetPartById(relationshipId) as ImagePart : null;
        }

        /// <summary>
        ///     Replaces the picture's underlying image with the provided stream.
        /// </summary>
        /// <param name="newImage">Stream containing the new image data.</param>
        /// <param name="type">Image format of the new image.</param>
        public void UpdateImage(Stream newImage, ImagePartType type) {
            if (newImage == null) {
                throw new ArgumentNullException(nameof(newImage));
            }

            PartTypeInfo partTypeInfo = type.ToPartTypeInfo();

            Picture picture = (Picture)Element;
            A.Blip blip = picture.BlipFill?.Blip ?? throw new InvalidOperationException("Picture has no image");

            if (blip.Embed != null) {
                OpenXmlPart? oldPart = _slidePart.GetPartById(blip.Embed!);
                if (oldPart != null) {
                    _slidePart.DeletePart(oldPart);
                }
            }

            string imageExtension = PowerPointPartFactory.GetImageExtension(type);
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
            newImage.Position = 0;
            imagePart.FeedData(newImage);
            string relId = _slidePart.GetIdOfPart(imagePart);
            blip.Embed = relId;
        }

        /// <summary>
        ///     Replaces the picture's underlying image with the provided file.
        /// </summary>
        /// <param name="imagePath">Path to the new image file.</param>
        public void UpdateImage(string imagePath) {
            if (imagePath == null) {
                throw new ArgumentNullException(nameof(imagePath));
            }
            if (!File.Exists(imagePath)) {
                throw new FileNotFoundException("Image file not found.", imagePath);
            }

            ImagePartType type = GetImagePartType(imagePath);
            using FileStream stream = new(imagePath, FileMode.Open, FileAccess.Read, FileShare.Read);
            UpdateImage(stream, type);
        }

        /// <summary>
        ///     Crops the image by the specified percentages (0-100).
        /// </summary>
        public void Crop(double leftPercent, double topPercent, double rightPercent, double bottomPercent) {
            ValidatePercent(leftPercent, nameof(leftPercent));
            ValidatePercent(topPercent, nameof(topPercent));
            ValidatePercent(rightPercent, nameof(rightPercent));
            ValidatePercent(bottomPercent, nameof(bottomPercent));

            if (IsZero(leftPercent) && IsZero(topPercent) && IsZero(rightPercent) && IsZero(bottomPercent)) {
                ResetCrop();
                return;
            }

            int left = ToCropValue(leftPercent);
            int top = ToCropValue(topPercent);
            int right = ToCropValue(rightPercent);
            int bottom = ToCropValue(bottomPercent);
            SetSourceRectangle(left, top, right, bottom);
        }

        /// <summary>
        ///     Removes any cropping from the image.
        /// </summary>
        public void ResetCrop() {
            SetSourceRectangle(null, null, null, null);
        }

        /// <summary>
        ///     Fits the image into the current shape bounds, optionally cropping to fill.
        /// </summary>
        /// <param name="imageWidth">Source image width (pixels).</param>
        /// <param name="imageHeight">Source image height (pixels).</param>
        /// <param name="crop">When true, crops to fill the box. When false, resizes to fit.</param>
        public void FitToBox(double imageWidth, double imageHeight, bool crop = true) {
            if (imageWidth <= 0) {
                throw new ArgumentOutOfRangeException(nameof(imageWidth));
            }
            if (imageHeight <= 0) {
                throw new ArgumentOutOfRangeException(nameof(imageHeight));
            }

            double boxWidth = Width;
            double boxHeight = Height;
            if (boxWidth <= 0 || boxHeight <= 0) {
                throw new InvalidOperationException("Picture has invalid dimensions.");
            }

            double imageAspect = imageWidth / imageHeight;
            double boxAspect = boxWidth / boxHeight;

            if (crop) {
                double left = 0;
                double top = 0;
                double right = 0;
                double bottom = 0;

                if (imageAspect > boxAspect) {
                    double cropRatio = 1 - (boxAspect / imageAspect);
                    left = cropRatio / 2;
                    right = cropRatio / 2;
                } else if (imageAspect < boxAspect) {
                    double cropRatio = 1 - (imageAspect / boxAspect);
                    top = cropRatio / 2;
                    bottom = cropRatio / 2;
                }

                Crop(left * 100, top * 100, right * 100, bottom * 100);
            } else {
                ResetCrop();

                double newWidth;
                double newHeight;

                if (imageAspect > boxAspect) {
                    newWidth = boxWidth;
                    newHeight = boxWidth / imageAspect;
                } else {
                    newHeight = boxHeight;
                    newWidth = boxHeight * imageAspect;
                }

                long deltaX = (long)Math.Round((boxWidth - newWidth) / 2);
                long deltaY = (long)Math.Round((boxHeight - newHeight) / 2);

                Left += deltaX;
                Top += deltaY;
                Width = (long)Math.Round(newWidth);
                Height = (long)Math.Round(newHeight);
            }
        }

        private static ImagePartType GetImagePartType(string imagePath) {
            string extension = Path.GetExtension(imagePath).ToLowerInvariant();
            return extension switch {
                ".jpg" or ".jpeg" => ImagePartType.Jpeg,
                ".gif" => ImagePartType.Gif,
                ".bmp" => ImagePartType.Bmp,
                _ => ImagePartType.Png
            };
        }

        private static void ValidatePercent(double value, string paramName) {
            if (value < 0 || value > 100) {
                throw new ArgumentOutOfRangeException(paramName, "Percent must be between 0 and 100.");
            }
        }

        private static int ToCropValue(double percent) {
            return (int)Math.Round(percent * 1000);
        }

        private static bool IsZero(double value) {
            return Math.Abs(value) < 0.000001d;
        }

        private void SetSourceRectangle(int? left, int? top, int? right, int? bottom) {
            A.SourceRectangle? rect = Picture.BlipFill?.SourceRectangle;

            if (left == null && top == null && right == null && bottom == null) {
                rect?.Remove();
                return;
            }

            if (Picture.BlipFill == null) {
                Picture.BlipFill = new BlipFill();
            }

            rect ??= new A.SourceRectangle();
            rect.Left = left;
            rect.Top = top;
            rect.Right = right;
            rect.Bottom = bottom;
            Picture.BlipFill.SourceRectangle = rect;
        }
    }
}
