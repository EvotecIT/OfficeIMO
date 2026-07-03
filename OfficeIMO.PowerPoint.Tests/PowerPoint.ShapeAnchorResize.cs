using System;
using System.IO;
using OfficeIMO.PowerPoint;
using Xunit;

namespace OfficeIMO.Tests {
    public class PowerPointShapeAnchorResizeTests {
        [Fact]
        public void Resize_WithCenterAnchor_PreservesCenter() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            try {
                using PowerPointPresentation presentation = PowerPointPresentation.Create(filePath);
                PowerPointSlide slide = presentation.AddSlide();
                PowerPointAutoShape shape = slide.AddRectangle(1000, 2000, 3000, 2000);

                long centerX = shape.CenterX;
                long centerY = shape.CenterY;

                shape.Resize(4000, 3000, PowerPointShapeAnchor.Center);

                Assert.Equal(centerX, shape.CenterX);
                Assert.Equal(centerY, shape.CenterY);
                Assert.Equal(4000, shape.Width);
                Assert.Equal(3000, shape.Height);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Resize_WithTopLeftAnchor_PreservesTopLeft() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            try {
                using PowerPointPresentation presentation = PowerPointPresentation.Create(filePath);
                PowerPointSlide slide = presentation.AddSlide();
                PowerPointAutoShape shape = slide.AddRectangle(1000, 2000, 3000, 2000);

                shape.Resize(5000, 4000, PowerPointShapeAnchor.TopLeft);

                Assert.Equal(1000, shape.Left);
                Assert.Equal(2000, shape.Top);
                Assert.Equal(5000, shape.Width);
                Assert.Equal(4000, shape.Height);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Scale_WithBottomRightAnchor_PreservesBottomRight() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            try {
                using PowerPointPresentation presentation = PowerPointPresentation.Create(filePath);
                PowerPointSlide slide = presentation.AddSlide();
                PowerPointAutoShape shape = slide.AddRectangle(1000, 2000, 3000, 2000);

                long right = shape.Right;
                long bottom = shape.Bottom;

                shape.Scale(0.5, PowerPointShapeAnchor.BottomRight);

                Assert.Equal(right, shape.Right);
                Assert.Equal(bottom, shape.Bottom);
                Assert.Equal(1500, shape.Width);
                Assert.Equal(1000, shape.Height);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }
    }
}
