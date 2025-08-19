using System;
using System.IO;
using OfficeIMO.PowerPoint;
using Xunit;

namespace OfficeIMO.Tests {
    public class PowerPointShapesManagement {
        [Fact]
        public void CanGetAndRemoveShapes() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            string imagePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Images", "BackgroundImage.png");

            using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                PowerPointSlide slide = presentation.AddSlide();
                PowerPointTextBox box1 = slide.AddTextBox("First");
                PowerPointPicture pic1 = slide.AddPicture(imagePath);
                PowerPointTextBox box2 = slide.AddTextBox("Second");
                PowerPointPicture pic2 = slide.AddPicture(imagePath);

                Assert.Equal("TextBox 1", box1.Name);
                Assert.Equal("TextBox 2", box2.Name);
                Assert.Equal("Picture 1", pic1.Name);
                Assert.Equal("Picture 2", pic2.Name);

                Assert.Same(box1, slide.GetTextBox("TextBox 1"));
                Assert.Same(pic2, slide.GetPicture("Picture 2"));
                Assert.Same(box1, slide.GetShape("TextBox 1"));

                slide.RemoveShape(pic1);
                Assert.Null(slide.GetPicture("Picture 1"));
                Assert.Equal(3, slide.Shapes.Count);

                presentation.Save();
            }

            File.Delete(filePath);
        }
    }
}
