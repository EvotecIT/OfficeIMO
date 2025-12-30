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
        public PowerPointPicture AddPicture(string imagePath, long left = 0L, long top = 0L, long width = 914400L,
            long height = 914400L) {
            if (imagePath == null) {
                throw new ArgumentNullException(nameof(imagePath));
            }

            if (!File.Exists(imagePath)) {
                throw new FileNotFoundException("Image file not found.", imagePath);
            }
            if (width <= 0) {
                throw new ArgumentOutOfRangeException(nameof(width));
            }
            if (height <= 0) {
                throw new ArgumentOutOfRangeException(nameof(height));
            }

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

            string name = GenerateUniqueName("Picture");
            DocumentFormat.OpenXml.Presentation.Picture picture = new(
                new NonVisualPictureProperties(
                    new NonVisualDrawingProperties { Id = _nextShapeId++, Name = name },
                    new NonVisualPictureDrawingProperties(new A.PictureLocks { NoChangeAspect = true }),
                    new ApplicationNonVisualDrawingProperties()
                ),
                new BlipFill(
                    new A.Blip { Embed = relationshipId },
                    new A.Stretch(new A.FillRectangle())
                ),
                new ShapeProperties(
                    new A.Transform2D(new A.Offset { X = left, Y = top }, new A.Extents { Cx = width, Cy = height }),
                    new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }
                )
            );

            CommonSlideData data = _slidePart.Slide.CommonSlideData ??= new CommonSlideData(new ShapeTree());
            ShapeTree tree = data.ShapeTree ??= new ShapeTree();
            tree.AppendChild(picture);
            PowerPointPicture pic = new(picture, _slidePart);
            _shapes.Add(pic);
            return pic;
        }

        /// <summary>
        ///     Adds an image from the provided stream.
        /// </summary>
        /// <param name="image">Stream containing the image data.</param>
        /// <param name="imageType">Image format of the stream.</param>
        /// <param name="left">Left position in EMUs.</param>
        /// <param name="top">Top position in EMUs.</param>
        /// <param name="width">Width in EMUs.</param>
        /// <param name="height">Height in EMUs.</param>
        public PowerPointPicture AddPicture(Stream image, ImagePartType imageType, long left = 0L, long top = 0L,
            long width = 914400L, long height = 914400L) {
            if (image == null) {
                throw new ArgumentNullException(nameof(image));
            }
            if (!image.CanRead) {
                throw new ArgumentException("Image stream must be readable.", nameof(image));
            }
            if (width <= 0) {
                throw new ArgumentOutOfRangeException(nameof(width));
            }
            if (height <= 0) {
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
                    new NonVisualDrawingProperties { Id = _nextShapeId++, Name = name },
                    new NonVisualPictureDrawingProperties(new A.PictureLocks { NoChangeAspect = true }),
                    new ApplicationNonVisualDrawingProperties()
                ),
                new BlipFill(
                    new A.Blip { Embed = relationshipId },
                    new A.Stretch(new A.FillRectangle())
                ),
                new ShapeProperties(
                    new A.Transform2D(new A.Offset { X = left, Y = top }, new A.Extents { Cx = width, Cy = height }),
                    new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }
                )
            );

            CommonSlideData data = _slidePart.Slide.CommonSlideData ??= new CommonSlideData(new ShapeTree());
            ShapeTree tree = data.ShapeTree ??= new ShapeTree();
            tree.AppendChild(picture);
            PowerPointPicture pic = new(picture, _slidePart);
            _shapes.Add(pic);
            return pic;
        }

        /// <summary>
        ///     Adds an image using a layout box.
        /// </summary>
        public PowerPointPicture AddPicture(string imagePath, PowerPointLayoutBox layout) {
            return AddPicture(imagePath, layout.Left, layout.Top, layout.Width, layout.Height);
        }

        /// <summary>
        ///     Adds an image from a stream using a layout box.
        /// </summary>
        public PowerPointPicture AddPicture(Stream image, ImagePartType imageType, PowerPointLayoutBox layout) {
            return AddPicture(image, imageType, layout.Left, layout.Top, layout.Width, layout.Height);
        }

        /// <summary>
        ///     Adds an image from the given file path using centimeter measurements.
        /// </summary>
        public PowerPointPicture AddPictureCm(string imagePath, double leftCm, double topCm, double widthCm,
            double heightCm) {
            return AddPicture(imagePath,
                PowerPointUnits.FromCentimeters(leftCm),
                PowerPointUnits.FromCentimeters(topCm),
                PowerPointUnits.FromCentimeters(widthCm),
                PowerPointUnits.FromCentimeters(heightCm));
        }

        /// <summary>
        ///     Adds an image from a stream using centimeter measurements.
        /// </summary>
        public PowerPointPicture AddPictureCm(Stream image, ImagePartType imageType, double leftCm, double topCm,
            double widthCm, double heightCm) {
            return AddPicture(image, imageType,
                PowerPointUnits.FromCentimeters(leftCm),
                PowerPointUnits.FromCentimeters(topCm),
                PowerPointUnits.FromCentimeters(widthCm),
                PowerPointUnits.FromCentimeters(heightCm));
        }

        /// <summary>
        ///     Adds an image from the given file path using inch measurements.
        /// </summary>
        public PowerPointPicture AddPictureInches(string imagePath, double leftInches, double topInches,
            double widthInches, double heightInches) {
            return AddPicture(imagePath,
                PowerPointUnits.FromInches(leftInches),
                PowerPointUnits.FromInches(topInches),
                PowerPointUnits.FromInches(widthInches),
                PowerPointUnits.FromInches(heightInches));
        }

        /// <summary>
        ///     Adds an image from a stream using inch measurements.
        /// </summary>
        public PowerPointPicture AddPictureInches(Stream image, ImagePartType imageType, double leftInches,
            double topInches, double widthInches, double heightInches) {
            return AddPicture(image, imageType,
                PowerPointUnits.FromInches(leftInches),
                PowerPointUnits.FromInches(topInches),
                PowerPointUnits.FromInches(widthInches),
                PowerPointUnits.FromInches(heightInches));
        }

        /// <summary>
        ///     Adds an image from the given file path using point measurements.
        /// </summary>
        public PowerPointPicture AddPicturePoints(string imagePath, double leftPoints, double topPoints,
            double widthPoints, double heightPoints) {
            return AddPicture(imagePath,
                PowerPointUnits.FromPoints(leftPoints),
                PowerPointUnits.FromPoints(topPoints),
                PowerPointUnits.FromPoints(widthPoints),
                PowerPointUnits.FromPoints(heightPoints));
        }

        /// <summary>
        ///     Adds an image from a stream using point measurements.
        /// </summary>
        public PowerPointPicture AddPicturePoints(Stream image, ImagePartType imageType, double leftPoints,
            double topPoints, double widthPoints, double heightPoints) {
            return AddPicture(image, imageType,
                PowerPointUnits.FromPoints(leftPoints),
                PowerPointUnits.FromPoints(topPoints),
                PowerPointUnits.FromPoints(widthPoints),
                PowerPointUnits.FromPoints(heightPoints));
        }

        private static ImagePartType GetImagePartType(string imagePath) {
            string extension = Path.GetExtension(imagePath).ToLowerInvariant();
            return extension switch {
                ".jpg" or ".jpeg" => ImagePartType.Jpeg,
                ".gif" => ImagePartType.Gif,
                ".bmp" => ImagePartType.Bmp,
                ".tif" or ".tiff" => ImagePartType.Tiff,
                ".emf" => ImagePartType.Emf,
                ".wmf" => ImagePartType.Wmf,
                ".ico" => ImagePartType.Icon,
                ".pcx" => ImagePartType.Pcx,
                _ => ImagePartType.Png
            };
        }

    }
}
