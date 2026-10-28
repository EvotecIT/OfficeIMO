using System;
using System.IO;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.PowerPoint {
    public partial class PowerPointSlide {
        /// <summary>
        ///     Adds an image from the given file path.
        /// </summary>
        public PowerPointPicture AddImage(string imagePath, OfficeLength? left = null, OfficeLength? top = null, OfficeLength? width = null,
            OfficeLength? height = null) {
            long nativeLeft = (left ?? OfficeLength.Emus(0L)).ToEmus();
            long nativeTop = (top ?? OfficeLength.Emus(0L)).ToEmus();
            long nativeWidth = (width ?? OfficeLength.Emus(914400L)).ToEmus();
            long nativeHeight = (height ?? OfficeLength.Emus(914400L)).ToEmus();
            if (imagePath == null) {
                throw new ArgumentNullException(nameof(imagePath));
            }

            if (!File.Exists(imagePath)) {
                throw new FileNotFoundException("Image file not found.", imagePath);
            }
            if (nativeWidth <= 0) {
                throw new ArgumentOutOfRangeException(nameof(width));
            }
            if (nativeHeight <= 0) {
                throw new ArgumentOutOfRangeException(nameof(height));
            }

            OfficeIMO.Drawing.OfficeImageFormat imageType = GetImagePartType(imagePath);
            PartTypeInfo partTypeInfo = imageType.ToPartTypeInfo();
            string imageExtension = PowerPointPartFactory.GetImageExtension(imageType, imagePath);
            string imagePartUri = PowerPointPartFactory.GetIndexedPartUri(
                _slidePart.OpenXmlPackage,
                "ppt/media",
                "image",
                imageExtension,
                allowBaseWithoutIndex: false);
            uint shapeId = AllocateShapeId();
            ImagePart imagePart = PowerPointPartFactory.CreatePart<ImagePart>(
                _slidePart,
                partTypeInfo.ContentType,
                imagePartUri);
            using FileStream stream = new(imagePath, FileMode.Open, FileAccess.Read);
            imagePart.FeedData(stream);
            string relationshipId = _slidePart.GetIdOfPart(imagePart);

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
            PowerPointPicture pic = TrackShape(new PowerPointPicture(picture, _slidePart));
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
        public PowerPointPicture AddImage(Stream image,OfficeIMO.Drawing.OfficeImageFormat imageType, OfficeLength? left = null, OfficeLength? top = null,
            OfficeLength? width = null, OfficeLength? height = null) {
            long nativeLeft = (left ?? OfficeLength.Emus(0L)).ToEmus();
            long nativeTop = (top ?? OfficeLength.Emus(0L)).ToEmus();
            long nativeWidth = (width ?? OfficeLength.Emus(914400L)).ToEmus();
            long nativeHeight = (height ?? OfficeLength.Emus(914400L)).ToEmus();
            if (image == null) {
                throw new ArgumentNullException(nameof(image));
            }
            if (!image.CanRead) {
                throw new ArgumentException("Image stream must be readable.", nameof(image));
            }
            if (nativeWidth <= 0) {
                throw new ArgumentOutOfRangeException(nameof(width));
            }
            if (nativeHeight <= 0) {
                throw new ArgumentOutOfRangeException(nameof(height));
            }

            PartTypeInfo partTypeInfo = imageType.ToPartTypeInfo();
            string imageExtension = PowerPointPartFactory.GetImageExtension(imageType);
            string imagePartUri = PowerPointPartFactory.GetIndexedPartUri(
                _slidePart.OpenXmlPackage,
                "ppt/media",
                "image",
                imageExtension,
                allowBaseWithoutIndex: false);
            uint shapeId = AllocateShapeId();
            ImagePart imagePart = PowerPointPartFactory.CreatePart<ImagePart>(
                _slidePart,
                partTypeInfo.ContentType,
                imagePartUri);
            if (image.CanSeek) {
                image.Position = 0;
            }
            imagePart.FeedData(image);
            string relationshipId = _slidePart.GetIdOfPart(imagePart);

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
            PowerPointPicture pic = TrackShape(new PowerPointPicture(picture, _slidePart));
            return pic;
        }

        /// <summary>
        ///     Adds an image using a layout box.
        /// </summary>
        public PowerPointPicture AddImage(string imagePath, PowerPointLayoutBox layout) {
            return AddImage(imagePath, layout.Left, layout.Top, layout.Width, layout.Height);
        }

        /// <summary>
        ///     Adds an image from a stream using a layout box.
        /// </summary>
        public PowerPointPicture AddImage(Stream image, OfficeIMO.Drawing.OfficeImageFormat imageType, PowerPointLayoutBox layout) {
            return AddImage(image, imageType, layout.Left, layout.Top, layout.Width, layout.Height);
        }

        private static OfficeIMO.Drawing.OfficeImageFormat GetImagePartType(string imagePath) {
            return PowerPointImageFormatExtensions.FromImagePath(imagePath);
        }

    }
}
