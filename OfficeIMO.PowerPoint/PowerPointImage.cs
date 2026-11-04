using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Office2019.Drawing.SVG;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.PowerPoint {
    /// <summary>
    ///     Represents an image placed on a slide.
    /// </summary>
    public partial class PowerPointImage : PowerPointShape {
        private readonly SlidePart? _initialSlidePart;
        private SlidePart? _slidePart => Element.Ancestors<Slide>().FirstOrDefault()?.SlidePart ?? _initialSlidePart;
        private readonly OpenXmlPartContainer _initialOwnerPart;
        private OpenXmlPartContainer _ownerPart => _initialOwnerPart is SlidePart ? _slidePart! : _initialOwnerPart;

        internal PowerPointImage(Picture picture, SlidePart slidePart) : this(picture, slidePart, slidePart) {
        }

        internal PowerPointImage(Picture picture, OpenXmlPartContainer ownerPart) : this(picture, null, ownerPart) {
        }

        private PowerPointImage(Picture picture, SlidePart? slidePart, OpenXmlPartContainer ownerPart) : base(picture) {
            _initialSlidePart = slidePart;
            _initialOwnerPart = ownerPart ?? throw new ArgumentNullException(nameof(ownerPart));
        }

        internal SlidePart SlidePart => _slidePart ?? throw new InvalidOperationException("This picture is not owned by a slide part.");

        private Picture Picture => (Picture)Element;

        /// <summary>
        ///     Gets the MIME content type of the underlying image.
        /// </summary>
        public string? ContentType => GetImagePart()?.ContentType;

        /// <summary>
        ///     Gets the MIME content type of the underlying image.
        /// </summary>
        public string? MimeType => ContentType;

        internal Uri? ClickHyperlink {
            get => PowerPointHyperlinkResolver.Resolve(_ownerPart, _slidePart ?? OwnerSlide?.SlidePart,
                Picture.NonVisualPictureProperties?.NonVisualDrawingProperties?.GetFirstChild<A.HyperlinkOnClick>());
        }

        private ImagePart? GetImagePart() {
            Picture picture = (Picture)Element;
            string? relationshipId = picture.BlipFill?.Blip?.Embed?.Value;
            return relationshipId != null && _ownerPart.TryGetPartById(relationshipId, out OpenXmlPart? part)
                ? part as ImagePart
                : null;
        }

        /// <summary>
        ///     Replaces the picture's underlying image with the provided stream.
        /// </summary>
        /// <param name="newImage">Stream containing the new image data.</param>
        /// <param name="type">Image format of the new image.</param>
        public void UpdateImage(Stream newImage, OfficeIMO.Drawing.OfficeImageFormat type) {
            if (newImage == null) {
                throw new ArgumentNullException(nameof(newImage));
            }

            using MemoryStream? normalized = PowerPointImageNormalization.NormalizeHeicForEmbedding(newImage, type);
            if (normalized != null) {
                UpdateImage(normalized, OfficeIMO.Drawing.OfficeImageFormat.Png);
                return;
            }

            PartTypeInfo partTypeInfo = type.ToPartTypeInfo();
            SlidePart slidePart = SlidePart;

            Picture picture = (Picture)Element;
            A.Blip blip = picture.BlipFill?.Blip ?? throw new InvalidOperationException("Picture has no image");
            string? previousRelationshipId = blip.Embed?.Value;

            string imageExtension = PowerPointPartFactory.GetImageExtension(type);
            string imagePartUri = PowerPointPartFactory.GetIndexedPartUri(
                slidePart.OpenXmlPackage,
                "ppt/media",
                "image",
                imageExtension,
                allowBaseWithoutIndex: false);
            ImagePart imagePart = PowerPointPartFactory.CreatePart<ImagePart>(
                slidePart,
                partTypeInfo.ContentType,
                imagePartUri);
            if (newImage.CanSeek) {
                newImage.Position = 0;
            }
            imagePart.FeedData(newImage);
            string relId = slidePart.GetIdOfPart(imagePart);
            blip.Embed = relId;
            string[] replacedVectorIds = RemoveVectorSources(blip);

            foreach (string relationshipId in replacedVectorIds.Prepend(previousRelationshipId).OfType<string>().Distinct(StringComparer.Ordinal)) {
                if (IsRelationshipReferenced(relationshipId)) continue;
                try {
                    OpenXmlPart? oldPart = slidePart.GetPartById(relationshipId);
                    if (oldPart != null) {
                        slidePart.DeletePart(oldPart);
                    }
                } catch (ArgumentOutOfRangeException) {
                    // The previous relationship may already be absent on damaged input.
                }
            }
        }

        /// <summary>
        ///     An SVG source in the blip extension list is drawn in place of the embedded raster, so it would keep
        ///     showing the old image after a replacement.
        /// </summary>
        private static string[] RemoveVectorSources(A.Blip blip) {
            A.BlipExtensionList? extensions = blip.GetFirstChild<A.BlipExtensionList>();
            if (extensions == null) return Array.Empty<string>();
            var removed = new List<string>();
            foreach (A.BlipExtension extension in extensions.Elements<A.BlipExtension>().ToArray()) {
                SVGBlip? vector = extension.GetFirstChild<SVGBlip>();
                if (vector == null) continue;
                if (vector.Embed?.Value is string relationshipId) removed.Add(relationshipId);
                extension.Remove();
            }
            if (!extensions.HasChildren) extensions.Remove();
            return removed.ToArray();
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

            OfficeIMO.Drawing.OfficeImageFormat type = GetImagePartType(imagePath);
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
        ///     Gets the left crop ratio authored on the picture, where 0 means no crop and 1 means the full width.
        /// </summary>
        public double CropLeftRatio => GetCrop().Left;

        /// <summary>
        ///     Gets the top crop ratio authored on the picture, where 0 means no crop and 1 means the full height.
        /// </summary>
        public double CropTopRatio => GetCrop().Top;

        /// <summary>
        ///     Gets the right crop ratio authored on the picture, where 0 means no crop and 1 means the full width.
        /// </summary>
        public double CropRightRatio => GetCrop().Right;

        /// <summary>
        ///     Gets the bottom crop ratio authored on the picture, where 0 means no crop and 1 means the full height.
        /// </summary>
        public double CropBottomRatio => GetCrop().Bottom;

        internal PowerPointImageCrop GetCrop() {
            A.SourceRectangle? rect = Picture.BlipFill?.SourceRectangle;
            if (rect == null) {
                return PowerPointImageCrop.None;
            }

            return new PowerPointImageCrop(
                ToCropFraction(rect.Left?.Value),
                ToCropFraction(rect.Top?.Value),
                ToCropFraction(rect.Right?.Value),
                ToCropFraction(rect.Bottom?.Value));
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

            double boxWidth = Width.ToEmus();
            double boxHeight = Height.ToEmus();
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

                Left += OfficeIMO.OfficeLength.Emus(deltaX);
                Top += OfficeIMO.OfficeLength.Emus(deltaY);
                Width = OfficeIMO.OfficeLength.Emus((long)Math.Round(newWidth));
                Height = OfficeIMO.OfficeLength.Emus((long)Math.Round(newHeight));
            }
        }

        private static OfficeIMO.Drawing.OfficeImageFormat GetImagePartType(string imagePath) {
            return PowerPointImageFormatExtensions.FromImagePath(imagePath);
        }

        private static void ValidatePercent(double value, string paramName) {
            if (value < 0 || value > 100) {
                throw new ArgumentOutOfRangeException(paramName, "Percent must be between 0 and 100.");
            }
        }

        private static int ToCropValue(double percent) {
            return (int)Math.Round(percent * 1000);
        }

        private static double ToCropFraction(int? value) {
            if (!value.HasValue) {
                return 0D;
            }

            return Math.Min(0.999999D, Math.Max(0D, value.Value / 100000D));
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

        private bool IsRelationshipReferenced(string relationshipId) {
            return SlideRoot.Descendants().Any(element => element is A.Blip { } blip && blip.Embed?.Value == relationshipId
                || element is SVGBlip { } vector && vector.Embed?.Value == relationshipId);
        }
    }
}
