using System;
using System.IO;
using System.Linq;
using OfficeIMO.PowerPoint;
using Xunit;

namespace OfficeIMO.Tests {
    public class PowerPointPictureStream {
        [Fact]
        public void CanAddPictureFromStream() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            string imagePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Images", "BackgroundImage.png");

            try {
                using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                    PowerPointSlide slide = presentation.AddSlide();
                    using FileStream stream = new(imagePath, FileMode.Open, FileAccess.Read, FileShare.Read);
                    PowerPointPicture picture = slide.AddPicture(stream, ImagePartType.Png, 0, 0, 1000000, 1000000);
                    Assert.Equal("image/png", picture.ContentType);
                    presentation.Save();
                }

                using (PowerPointPresentation presentation = PowerPointPresentation.Open(filePath)) {
                    PowerPointSlide slide = presentation.Slides.Single();
                    Assert.Single(slide.Pictures);
                }
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }
    }
}
