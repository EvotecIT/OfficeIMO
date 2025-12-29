using System;
using System.IO;
using OfficeIMO.PowerPoint;
using Xunit;

namespace OfficeIMO.Tests {
    public class PowerPointShapeResizeTests {
        [Fact]
        public void ResizeShapes_Width_UsesLargest() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            try {
                using PowerPointPresentation presentation = PowerPointPresentation.Create(filePath);
                PowerPointSlide slide = presentation.AddSlide();
                PowerPointAutoShape a = slide.AddRectangle(0, 0, 1000, 2000);
                PowerPointAutoShape b = slide.AddRectangle(0, 0, 2000, 3000);
                PowerPointAutoShape c = slide.AddRectangle(0, 0, 1500, 4000);

                slide.ResizeShapes(slide.Shapes, PowerPointShapeSizeDimension.Width, PowerPointShapeSizeReference.Largest);

                Assert.Equal(2000, a.Width);
                Assert.Equal(2000, b.Width);
                Assert.Equal(2000, c.Width);
                Assert.Equal(2000, a.Height);
                Assert.Equal(3000, b.Height);
                Assert.Equal(4000, c.Height);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void ResizeShapes_Both_UsesAverage() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            try {
                using PowerPointPresentation presentation = PowerPointPresentation.Create(filePath);
                PowerPointSlide slide = presentation.AddSlide();
                PowerPointAutoShape a = slide.AddRectangle(0, 0, 1000, 2000);
                PowerPointAutoShape b = slide.AddRectangle(0, 0, 3000, 4000);

                slide.ResizeShapes(slide.Shapes, PowerPointShapeSizeDimension.Both, PowerPointShapeSizeReference.Average);

                Assert.Equal(2000, a.Width);
                Assert.Equal(2000, b.Width);
                Assert.Equal(3000, a.Height);
                Assert.Equal(3000, b.Height);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void ResizeShapes_Height_UsesFirst() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            try {
                using PowerPointPresentation presentation = PowerPointPresentation.Create(filePath);
                PowerPointSlide slide = presentation.AddSlide();
                PowerPointAutoShape a = slide.AddRectangle(0, 0, 1000, 500);
                PowerPointAutoShape b = slide.AddRectangle(0, 0, 1000, 900);

                slide.ResizeShapes(slide.Shapes, PowerPointShapeSizeDimension.Height, PowerPointShapeSizeReference.First);

                Assert.Equal(500, a.Height);
                Assert.Equal(500, b.Height);
                Assert.Equal(1000, a.Width);
                Assert.Equal(1000, b.Width);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }
    }
}
