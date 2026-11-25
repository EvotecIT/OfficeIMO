using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.PowerPoint {
    public partial class PowerPointSlide {
        /// <summary>
        ///     Adds an image from the given file path.
        /// </summary>
#pragma warning disable RS0026 // Existing overload shape; only the return type was renamed to PowerPointImage.
        public PowerPointImage AddImage(string imagePath, OfficeLength? left = null, OfficeLength? top = null, OfficeLength? width = null,
            OfficeLength? height = null) {
#pragma warning restore RS0026
            long nativeLeft = (left ?? OfficeLength.Emus(0L)).ToEmus();
            long nativeTop = (top ?? OfficeLength.Emus(0L)).ToEmus();
            long nativeWidth = (width ?? OfficeLength.Emus(914400L)).ToEmus();
            long nativeHeight = (height ?? OfficeLength.Emus(914400L)).ToEmus();
            ArgumentNullException.ThrowIfNull(imagePath);

            if (!File.Exists(imagePath)) {
                throw new FileNotFoundException("Image file not found.", imagePath);
            }
            ArgumentOutOfRangeException.ThrowIfNegativeOrZero(nativeWidth, nameof(width));
            ArgumentOutOfRangeException.ThrowIfNegativeOrZero(nativeHeight, nameof(height));

            using FileStream source = new(imagePath, FileMode.Open, FileAccess.Read);
            using MemoryStream? normalized = PowerPointImageNormalization.NormalizeHeicForEmbedding(source, OfficeIMO.Drawing.OfficeImageReader.FromExtension(imagePath));
            if (normalized != null) return AddImage(normalized, OfficeIMO.Drawing.OfficeImageFormat.Png, left, top, width, height);

            OfficeIMO.Drawing.OfficeImageFormat imageType = GetImagePartType(imagePath);
            if (imageType == OfficeIMO.Drawing.OfficeImageFormat.Svg) return AddImage(source, imageType, left, top, width, height);
            string imageExtension = PowerPointPartFactory.GetImageExtension(imageType, imagePath);
            uint shapeId = AllocateShapeId();
            (string relationshipId, int pixelWidth, int pixelHeight) = AddSlideImagePart(
                OfficeIMO.Core.Internal.OfficeStreamReader.ReadAllBytes(source), imageType, imageExtension);

            string name = GenerateUniqueName("Picture");
            DocumentFormat.OpenXml.Presentation.Picture picture = new(
                new NonVisualPictureProperties(
                    new NonVisualDrawingProperties { Id = shapeId, Name = name },
                    new NonVisualPictureDrawingProperties(new A.PictureLocks { NoChangeAspect = true }),
                    new ApplicationNonVisualDrawingProperties()
                ),
                new BlipFill(
                    new A.Blip { Embed = relationshipId },
                    new A.Stretch(new A.FillRectangle())
                ),
                new ShapeProperties(
                    new A.Transform2D(new A.Offset { X = nativeLeft, Y = nativeTop }, new A.Extents { Cx = nativeWidth, Cy = nativeHeight }),
                    new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }
                )
            );

            CommonSlideData data = SlideRoot.CommonSlideData ??= new CommonSlideData(new ShapeTree());
            ShapeTree tree = data.ShapeTree ??= new ShapeTree();
            tree.AppendChild(picture);
            PowerPointImage pic = TrackShape(new PowerPointImage(picture, SlidePart));
            // Keep the image's own aspect ratio inside the requested box instead of stretching it.
            pic.ContainInFrame(pixelWidth, pixelHeight);
            return pic;
        }

        /// <summary>
        ///     Adds an image from the provided stream.
        /// </summary>
        /// <param name="image">Stream containing the image data.</param>
        /// <param name="imageType">Image format of the stream.</param>
        /// <param name="left">Left position as a physical length.</param>
        /// <param name="top">Top position as a physical length.</param>
        /// <param name="width">Width as a physical length.</param>
        /// <param name="height">Height as a physical length.</param>
#pragma warning disable RS0026 // Existing overload shape; only the return type was renamed to PowerPointImage.
        public PowerPointImage AddImage(Stream image,OfficeIMO.Drawing.OfficeImageFormat imageType, OfficeLength? left = null, OfficeLength? top = null,
            OfficeLength? width = null, OfficeLength? height = null) {
#pragma warning restore RS0026
            long nativeLeft = (left ?? OfficeLength.Emus(0L)).ToEmus();
            long nativeTop = (top ?? OfficeLength.Emus(0L)).ToEmus();
            long nativeWidth = (width ?? OfficeLength.Emus(914400L)).ToEmus();
            long nativeHeight = (height ?? OfficeLength.Emus(914400L)).ToEmus();
            ArgumentNullException.ThrowIfNull(image);
            if (!image.CanRead) {
                throw new ArgumentException("Image stream must be readable.", nameof(image));
            }
            ArgumentOutOfRangeException.ThrowIfNegativeOrZero(nativeWidth, nameof(width));
            ArgumentOutOfRangeException.ThrowIfNegativeOrZero(nativeHeight, nameof(height));

            using MemoryStream? normalized = PowerPointImageNormalization.NormalizeHeicForEmbedding(image, imageType);
            if (normalized != null) return AddImage(normalized, OfficeIMO.Drawing.OfficeImageFormat.Png, left, top, width, height);

            if (image.CanSeek) {
                image.Position = 0;
            }
            // An SVG picture stores its PNG fallback in the ordinary blip and the SVG in the Office 2016 blip extension.
            byte[] imageBytes = OfficeIMO.Core.Internal.OfficeStreamReader.ReadAllBytes(image);
            uint shapeId = AllocateShapeId();
            string relationshipId;
            int pixelWidth, pixelHeight;
            A.Blip blip;
            if (imageType == OfficeIMO.Drawing.OfficeImageFormat.Svg
                && TryAddSvgImageParts(imageBytes, out relationshipId, out string svgRelationshipId, out pixelWidth, out pixelHeight)) {
                blip = new A.Blip { Embed = relationshipId };
                blip.Append(new A.BlipExtensionList(new A.BlipExtension(
                    new DocumentFormat.OpenXml.Office2019.Drawing.SVG.SVGBlip { Embed = svgRelationshipId }) {
                    Uri = OfficeIMO.Core.Internal.OfficeSvgPictureFallback.BlipExtensionUri
                }));
            } else {
                (relationshipId, pixelWidth, pixelHeight) = AddSlideImagePart(imageBytes, imageType);
                blip = new A.Blip { Embed = relationshipId };
            }

            string name = GenerateUniqueName("Picture");
            DocumentFormat.OpenXml.Presentation.Picture picture = new(
                new NonVisualPictureProperties(
                    new NonVisualDrawingProperties { Id = shapeId, Name = name },
                    new NonVisualPictureDrawingProperties(new A.PictureLocks { NoChangeAspect = true }),
                    new ApplicationNonVisualDrawingProperties()
                ),
                new BlipFill(
                    blip,
                    new A.Stretch(new A.FillRectangle())
                ),
                new ShapeProperties(
                    new A.Transform2D(new A.Offset { X = nativeLeft, Y = nativeTop }, new A.Extents { Cx = nativeWidth, Cy = nativeHeight }),
                    new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }
                )
            );

            CommonSlideData data = SlideRoot.CommonSlideData ??= new CommonSlideData(new ShapeTree());
            ShapeTree tree = data.ShapeTree ??= new ShapeTree();
            tree.AppendChild(picture);
            PowerPointImage pic = TrackShape(new PowerPointImage(picture, SlidePart));
            // Keep the image's own aspect ratio inside the requested box instead of stretching it.
            pic.ContainInFrame(pixelWidth, pixelHeight);
            return pic;
        }

        /// <summary>
        ///     Adds an image using a layout box.
        /// </summary>
        public PowerPointImage AddImage(string imagePath, PowerPointLayoutBox layout) {
            return AddImage(imagePath, layout.Left, layout.Top, layout.Width, layout.Height);
        }

        /// <summary>
        ///     Adds an image from a stream using a layout box.
        /// </summary>
        public PowerPointImage AddImage(Stream image, OfficeIMO.Drawing.OfficeImageFormat imageType, PowerPointLayoutBox layout) {
            return AddImage(image, imageType, layout.Left, layout.Top, layout.Width, layout.Height);
        }

        /// <summary>
        ///     Adds an SVG picture's PNG fallback and SVG source. The fallback is rasterized once per distinct SVG in
        ///     the document; later additions of the same SVG reuse both parts.
        /// </summary>
        private bool TryAddSvgImageParts(byte[] svg, out string fallbackRelationshipId, out string svgRelationshipId,
            out int pixelWidth, out int pixelHeight) {
            PowerPointDocument? document = OwnerDocument;
            OpenXmlPackage package = SlidePart.OpenXmlPackage;
            string fallbackKey = PowerPointDocument.CreateImageContentKey("svg-fallback", svg);
            if (document != null && document.TryGetReusableImagePart(fallbackKey, package, out ReusableImagePart fallback)) {
                fallbackRelationshipId = LinkImagePart(fallback.Part);
                pixelWidth = fallback.Width;
                pixelHeight = fallback.Height;
            } else {
                if (!OfficeIMO.Core.Internal.OfficeSvgPictureFallback.TryCreate(svg, out byte[] png, out _, out _)) {
                    fallbackRelationshipId = svgRelationshipId = string.Empty;
                    pixelWidth = pixelHeight = 0;
                    return false;
                }
                (fallbackRelationshipId, pixelWidth, pixelHeight) = AddSlideImagePart(png,
                    OfficeIMO.Drawing.OfficeImageFormat.Png, contentKey: fallbackKey);
            }
            (svgRelationshipId, _, _) = AddSlideImagePart(svg, OfficeIMO.Drawing.OfficeImageFormat.Svg);
            return true;
        }

        /// <summary>
        ///     Links the slide to an image part holding <paramref name="data" />, creating the part only when the
        ///     document has none with the same content yet.
        /// </summary>
        private (string RelationshipId, int PixelWidth, int PixelHeight) AddSlideImagePart(byte[] data,
            OfficeIMO.Drawing.OfficeImageFormat imageType, string? imageExtension = null, string? contentKey = null) {
            PartTypeInfo partTypeInfo = imageType.ToPartTypeInfo();
            PowerPointDocument? document = OwnerDocument;
            OpenXmlPackage package = SlidePart.OpenXmlPackage;
            string key = contentKey ?? PowerPointDocument.CreateImageContentKey(partTypeInfo.ContentType, data);
            if (document != null && document.TryGetReusableImagePart(key, package, out ReusableImagePart reusable)) {
                return (LinkImagePart(reusable.Part), reusable.Width, reusable.Height);
            }

            string imagePartUri = PowerPointPartFactory.GetIndexedPartUri(
                package,
                "ppt/media",
                "image",
                imageExtension ?? PowerPointPartFactory.GetImageExtension(imageType),
                allowBaseWithoutIndex: false);
            ImagePart imagePart = PowerPointPartFactory.CreatePart<ImagePart>(
                SlidePart,
                partTypeInfo.ContentType,
                imagePartUri);
            using (var source = new MemoryStream(data, writable: false)) {
                imagePart.FeedData(source);
            }
            if (!PowerPointImageSize.TryRead(data, out int width, out int height)) {
                width = height = 0;
            }
            document?.RegisterReusableImagePart(key, new ReusableImagePart(imagePart, width, height));
            return (SlidePart.GetIdOfPart(imagePart), width, height);
        }

        private string LinkImagePart(ImagePart part) {
            foreach (IdPartPair pair in SlidePart.Parts) {
                if (ReferenceEquals(pair.OpenXmlPart, part)) return pair.RelationshipId;
            }
            return SlidePart.CreateRelationshipToPart(part);
        }

        private static OfficeIMO.Drawing.OfficeImageFormat GetImagePartType(string imagePath) {
            return PowerPointImageFormatExtensions.FromImagePath(imagePath);
        }

    }
}
