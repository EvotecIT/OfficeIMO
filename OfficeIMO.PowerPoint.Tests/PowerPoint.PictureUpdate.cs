using System;
using System.IO;
using System.Linq;
using PresentationDocument = DocumentFormat.OpenXml.Packaging.PresentationDocument;
using OfficeIMO.Drawing;
using OfficeIMO.PowerPoint;
using Xunit;

namespace OfficeIMO.Tests {
    public class PowerPointPictureUpdate {
        public static IEnumerable<object[]> ImageData => new[] {
            new object[] { "BackgroundImage.png", OfficeImageFormat.Png, "image/png" },
            new object[] { "Kulek.jpg", OfficeImageFormat.Jpeg, "image/jpeg" },
            new object[] { "example.gif", OfficeImageFormat.Gif, "image/gif" },
            new object[] { "snail.bmp", OfficeImageFormat.Bmp, "image/bmp" },
        };

        [Theory]
        [MemberData(nameof(ImageData))]
        public void CanAddPictureFromPathWithSharedImageExtensionMapping(string image, OfficeImageFormat expectedType, string expectedContentType) {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            string imagePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Images", image);
            Assert.Equal(expectedType, PowerPointImageFormatExtensions.FromImagePath(imagePath));

            using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                PowerPointSlide slide = presentation.AddSlide();
                PowerPointPicture picture = slide.AddPicture(imagePath);

                Assert.Equal(expectedContentType, picture.ContentType);
                Assert.Equal(expectedContentType, picture.MimeType);
                presentation.Save();
            }

            using (PowerPointPresentation presentation = PowerPointPresentation.Load(filePath)) {
                PowerPointPicture picture = Assert.Single(presentation.Slides.Single().Pictures);
                Assert.Equal(expectedContentType, picture.ContentType);
                Assert.Equal(expectedContentType, picture.MimeType);
            }

            File.Delete(filePath);
        }

        [Fact]
        public void PowerPointImagePartExtensionsUseSharedDrawingPolicy() {
            Assert.Equal(OfficeImageFormat.Png, OfficeImageFormat.Png.EnsurePowerPointImagePartSupport());
            Assert.Equal(".png", PowerPointPartFactory.GetImageExtension(OfficeImageFormat.Png));
            Assert.Equal(".jpeg", PowerPointPartFactory.GetImageExtension(OfficeImageFormat.Jpeg));
            Assert.Equal(".svg", PowerPointPartFactory.GetImageExtension(OfficeImageFormat.Svg));
            Assert.Equal(".emf", PowerPointPartFactory.GetImageExtension(OfficeImageFormat.Emf));
            Assert.Equal(".jpg", PowerPointPartFactory.GetImageExtension(OfficeImageFormat.Jpeg, @"C:\Temp\photo.JPG"));
        }

        [Fact]
        public void PowerPointImagePartExtensionsRejectUnsupportedWebP() {
            Assert.Throws<NotSupportedException>(() => OfficeImageFormat.Webp.EnsurePowerPointImagePartSupport());
        }

        [Fact]
        public void DanglingPictureRelationshipReturnsNoContentTypeInsteadOfThrowing() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            string imagePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Images", "BackgroundImage.png");
            try {
                using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                    presentation.AddSlide().AddPicture(imagePath);
                    presentation.Save();
                }
                using (PresentationDocument package = PresentationDocument.Open(filePath, true)) {
                    DocumentFormat.OpenXml.Presentation.Picture picture = package.PresentationPart!.SlideParts
                        .Single().Slide.Descendants<DocumentFormat.OpenXml.Presentation.Picture>().Single();
                    picture.BlipFill!.Blip!.Embed = "rIdMissingPreview";
                    package.PresentationPart.SlideParts.Single().Slide.Save();
                }

                using PowerPointPresentation reopened = PowerPointPresentation.Load(filePath);
                PowerPointPicture pictureModel = Assert.Single(reopened.Slides.Single().Pictures);
                Assert.Null(pictureModel.ContentType);
                Assert.Null(pictureModel.MimeType);
            } finally {
                if (File.Exists(filePath)) File.Delete(filePath);
            }
        }

        [Theory]
        [MemberData(nameof(ImageData))]
        public void CanUpdatePicture(string newImage, OfficeImageFormat type, string expectedContentType) {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            string originalImage = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Images", "BackgroundImage.png");
            string newImagePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Images", newImage);

            using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                PowerPointSlide slide = presentation.AddSlide();
                PowerPointPicture picture = slide.AddPicture(originalImage);
                using FileStream stream = new(newImagePath, FileMode.Open, FileAccess.Read);
                picture.UpdateImage(stream, type);
                Assert.Equal(expectedContentType, picture.ContentType);
                Assert.Equal(expectedContentType, picture.MimeType);
                presentation.Save();
            }

            using (PowerPointPresentation presentation = PowerPointPresentation.Load(filePath)) {
                PowerPointSlide slide = presentation.Slides.Single();
                PowerPointPicture picture = slide.Pictures.First();
                Assert.Equal(expectedContentType, picture.ContentType);
                Assert.Equal(expectedContentType, picture.MimeType);
            }

            File.Delete(filePath);
        }

        [Fact]
        public void UpdatingDuplicatedPictureDoesNotBreakSiblingRelationship() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            string originalImage = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Images", "BackgroundImage.png");
            string replacementImage = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Images", "Kulek.jpg");

            using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                PowerPointSlide slide = presentation.AddSlide();
                PowerPointPicture original = slide.AddPicture(originalImage);
                PowerPointPicture duplicate = Assert.IsType<PowerPointPicture>(slide.DuplicateShape(original, 250000, 0));

                using FileStream stream = new(replacementImage, FileMode.Open, FileAccess.Read);
                original.UpdateImage(stream, OfficeImageFormat.Jpeg);

                Assert.Equal("image/jpeg", original.ContentType);
                Assert.Equal("image/png", duplicate.ContentType);
                presentation.Save();
            }

            using (PowerPointPresentation presentation = PowerPointPresentation.Load(filePath)) {
                PowerPointPicture[] pictures = presentation.Slides.Single().Pictures.ToArray();
                Assert.Equal(2, pictures.Length);
                Assert.Contains(pictures, picture => picture.ContentType == "image/jpeg");
                Assert.Contains(pictures, picture => picture.ContentType == "image/png");
            }

            File.Delete(filePath);
        }
    }
}
