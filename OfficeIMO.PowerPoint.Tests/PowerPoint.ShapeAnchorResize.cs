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
                using PowerPointDocument presentation = PowerPointDocument.Create(filePath);
                PowerPointSlide slide = presentation.AddSlide();
                PowerPointAutoShape shape = slide.AddRectangle(1000, 2000, 3000, 2000);

                long centerX = shape.CenterX.ToEmus();
                long centerY = shape.CenterY.ToEmus();

                shape.Resize(OfficeIMO.OfficeLength.Emus(4000),OfficeIMO.OfficeLength.Emus(3000),PowerPointShapeAnchor.Center);

                Assert.Equal(centerX, shape.CenterX.ToEmus());
                Assert.Equal(centerY, shape.CenterY.ToEmus());
                Assert.Equal(4000, shape.Width.ToEmus());
                Assert.Equal(3000, shape.Height.ToEmus());
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
                using PowerPointDocument presentation = PowerPointDocument.Create(filePath);
                PowerPointSlide slide = presentation.AddSlide();
                PowerPointAutoShape shape = slide.AddRectangle(1000, 2000, 3000, 2000);

                shape.Resize(OfficeIMO.OfficeLength.Emus(5000),OfficeIMO.OfficeLength.Emus(4000),PowerPointShapeAnchor.TopLeft);

                Assert.Equal(1000, shape.Left.ToEmus());
                Assert.Equal(2000, shape.Top.ToEmus());
                Assert.Equal(5000, shape.Width.ToEmus());
                Assert.Equal(4000, shape.Height.ToEmus());
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
                using PowerPointDocument presentation = PowerPointDocument.Create(filePath);
                PowerPointSlide slide = presentation.AddSlide();
                PowerPointAutoShape shape = slide.AddRectangle(1000, 2000, 3000, 2000);

                long right = shape.Right.ToEmus();
                long bottom = shape.Bottom.ToEmus();

                shape.Scale(0.5, PowerPointShapeAnchor.BottomRight);

                Assert.Equal(right, shape.Right.ToEmus());
                Assert.Equal(bottom, shape.Bottom.ToEmus());
                Assert.Equal(1500, shape.Width.ToEmus());
                Assert.Equal(1000, shape.Height.ToEmus());
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }
    }
}
