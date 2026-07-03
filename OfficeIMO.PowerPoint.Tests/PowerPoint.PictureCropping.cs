using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.PowerPoint;
using Xunit;

namespace OfficeIMO.Tests {
    public class PowerPointPictureCropping {
        [Fact]
        public void CanApplyCropToPicture() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            string imagePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Images", "BackgroundImage.png");

            using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                PowerPointSlide slide = presentation.AddSlide();
                PowerPointPicture picture = slide.AddPicture(imagePath, 0, 0, 2000000, 1000000);
                picture.Crop(10, 5, 10, 0);
                presentation.Save();
            }

            using (PresentationDocument document = PresentationDocument.Open(filePath, false)) {
                var pic = document.PresentationPart!.SlideParts.First().Slide
                    .Descendants<DocumentFormat.OpenXml.Presentation.Picture>()
                    .First();
                var rect = pic.BlipFill?.SourceRectangle;
                Assert.NotNull(rect);
                Assert.Equal(10000, rect!.Left?.Value);
                Assert.Equal(5000, rect.Top?.Value);
                Assert.Equal(10000, rect.Right?.Value);
                Assert.Equal(0, rect.Bottom?.Value);
            }

            File.Delete(filePath);
        }

        [Fact]
        public void FitToBoxWithoutCropAdjustsSize() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            string imagePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Images", "BackgroundImage.png");

            using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                PowerPointSlide slide = presentation.AddSlide();
                PowerPointPicture picture = slide.AddPicture(imagePath, 0, 0, 2000000, 2000000);
                picture.FitToBox(400, 200, crop: false);
                presentation.Save();
            }

            using (PowerPointPresentation presentation = PowerPointPresentation.Open(filePath)) {
                PowerPointPicture picture = presentation.Slides.Single().Pictures.First();
                Assert.Equal(0, picture.Left);
                Assert.Equal(500000, picture.Top);
                Assert.Equal(2000000, picture.Width);
                Assert.Equal(1000000, picture.Height);
            }

            using (PresentationDocument document = PresentationDocument.Open(filePath, false)) {
                var pic = document.PresentationPart!.SlideParts.First().Slide
                    .Descendants<DocumentFormat.OpenXml.Presentation.Picture>()
                    .First();
                Assert.Null(pic.BlipFill?.SourceRectangle);
            }

            File.Delete(filePath);
        }
    }
}
