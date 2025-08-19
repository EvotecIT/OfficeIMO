using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.PowerPoint;
using Xunit;
using ImagePartType = DocumentFormat.OpenXml.Packaging.PartTypeInfo;

namespace OfficeIMO.Tests {
    public class PowerPointPictureUpdate {
        public static IEnumerable<object[]> ImageData => new[] {
            new object[] { "BackgroundImage.png", DocumentFormat.OpenXml.Packaging.ImagePartType.Png, "image/png" },
            new object[] { "Kulek.jpg", DocumentFormat.OpenXml.Packaging.ImagePartType.Jpeg, "image/jpeg" },
            new object[] { "example.gif", DocumentFormat.OpenXml.Packaging.ImagePartType.Gif, "image/gif" },
            new object[] { "snail.bmp", DocumentFormat.OpenXml.Packaging.ImagePartType.Bmp, "image/bmp" },
        };

        [Theory]
        [MemberData(nameof(ImageData))]
        public void CanUpdatePicture(string newImage, ImagePartType type, string expectedContentType) {
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

            using (PowerPointPresentation presentation = PowerPointPresentation.Open(filePath)) {
                PowerPointSlide slide = presentation.Slides.Single();
                PowerPointPicture picture = slide.Pictures.First();
                Assert.Equal(expectedContentType, picture.ContentType);
                Assert.Equal(expectedContentType, picture.MimeType);
            }

            File.Delete(filePath);
        }
    }
}
