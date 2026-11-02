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
                using (PowerPointDocument presentation = PowerPointDocument.Create(filePath)) {
                    PowerPointSlide slide = presentation.AddSlide();
                    using FileStream stream = new(imagePath, FileMode.Open, FileAccess.Read, FileShare.Read);
                    PowerPointPicture picture = slide.AddImage(stream, OfficeIMO.Drawing.OfficeImageFormat.Png, global::OfficeIMO.OfficeLength.Emus(0), global::OfficeIMO.OfficeLength.Emus(0), global::OfficeIMO.OfficeLength.Emus(1000000), global::OfficeIMO.OfficeLength.Emus(1000000));
                    Assert.Equal("image/png", picture.ContentType);
                    presentation.Save();
                }

                using (PowerPointDocument presentation = PowerPointDocument.Load(filePath)) {
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
