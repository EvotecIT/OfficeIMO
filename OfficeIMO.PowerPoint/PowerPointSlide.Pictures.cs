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

        private static ImagePartType GetImagePartType(string imagePath) {       
            string extension = Path.GetExtension(imagePath).ToLowerInvariant(); 
            return extension switch {
                ".jpg" or ".jpeg" => ImagePartType.Jpeg,
                ".gif" => ImagePartType.Gif,
                ".bmp" => ImagePartType.Bmp,
                _ => ImagePartType.Png
            };
        }

    }
}
