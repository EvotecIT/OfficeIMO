using System;
using System.IO;
using OfficeIMO.PowerPoint;
using Xunit;

namespace OfficeIMO.Tests {
    public class PowerPointShapeStackTests {
        [Fact]
        public void StackShapes_Horizontal_AlignsBottom() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            try {
                using PowerPointPresentation presentation = PowerPointPresentation.Create(filePath);
                PowerPointSlide slide = presentation.AddSlide();
                PowerPointAutoShape a = slide.AddRectangle(0, 0, 1000, 500);
                PowerPointAutoShape b = slide.AddRectangle(0, 0, 2000, 1000);

                var bounds = new PowerPointLayoutBox(100, 200, 8000, 3000);
                slide.StackShapes(slide.Shapes, PowerPointShapeStackDirection.Horizontal, bounds,
                    spacingEmus: 300, PowerPointShapeAlignment.Bottom);

                Assert.Equal(100, a.Left);
                Assert.Equal(100 + 1000 + 300, b.Left);
                Assert.Equal(bounds.Bottom - a.Height, a.Top);
                Assert.Equal(bounds.Bottom - b.Height, b.Top);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void StackShapes_Vertical_AlignsCenter() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            try {
                using PowerPointPresentation presentation = PowerPointPresentation.Create(filePath);
                PowerPointSlide slide = presentation.AddSlide();
                PowerPointAutoShape a = slide.AddRectangle(0, 0, 1000, 600);
                PowerPointAutoShape b = slide.AddRectangle(0, 0, 2000, 800);

                var bounds = new PowerPointLayoutBox(200, 300, 6000, 4000);
                slide.StackShapes(slide.Shapes, PowerPointShapeStackDirection.Vertical, bounds,
                    spacingEmus: 250, PowerPointShapeAlignment.Center);

                Assert.Equal(300, a.Top);
                Assert.Equal(300 + 600 + 250, b.Top);
                Assert.Equal(bounds.Left + (bounds.Width - a.Width) / 2, a.Left);
                Assert.Equal(bounds.Left + (bounds.Width - b.Width) / 2, b.Left);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void StackShapesToSlideContent_UsesMargin() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            try {
                using PowerPointPresentation presentation = PowerPointPresentation.Create(filePath);
                PowerPointSlide slide = presentation.AddSlide();
                PowerPointAutoShape a = slide.AddRectangle(0, 0, 1000, 700);
                PowerPointAutoShape b = slide.AddRectangle(0, 0, 1000, 900);
                long margin = PowerPointUnits.FromCentimeters(1);

                slide.StackShapesToSlideContent(slide.Shapes, PowerPointShapeStackDirection.Vertical,
                    spacingEmus: 200, marginEmus: margin);

                PowerPointLayoutBox content = presentation.SlideSize.GetContentBox(margin);
                Assert.Equal(content.Top, a.Top);
                Assert.Equal(content.Top + a.Height + 200, b.Top);
                Assert.Equal(content.Left, a.Left);
                Assert.Equal(content.Left, b.Left);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void StackShapes_Horizontal_JustifiesCenter() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            try {
                using PowerPointPresentation presentation = PowerPointPresentation.Create(filePath);
                PowerPointSlide slide = presentation.AddSlide();
                PowerPointAutoShape a = slide.AddRectangle(0, 0, 1000, 500);
                PowerPointAutoShape b = slide.AddRectangle(0, 0, 2000, 500);

                var bounds = new PowerPointLayoutBox(0, 0, 8000, 2000);
                slide.StackShapes(slide.Shapes, PowerPointShapeStackDirection.Horizontal, bounds,
                    spacingEmus: 500, PowerPointShapeStackJustify.Center);

                long totalWidth = a.Width + b.Width + 500;
                long expectedLeft = (bounds.Width - totalWidth) / 2;

                Assert.Equal(expectedLeft, a.Left);
                Assert.Equal(expectedLeft + a.Width + 500, b.Left);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void StackShapes_ClampSpacingToBounds() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            try {
                using PowerPointPresentation presentation = PowerPointPresentation.Create(filePath);
                PowerPointSlide slide = presentation.AddSlide();
                PowerPointAutoShape a = slide.AddRectangle(0, 0, 2500, 500);
                PowerPointAutoShape b = slide.AddRectangle(0, 0, 2500, 500);
                PowerPointAutoShape c = slide.AddRectangle(0, 0, 2500, 500);

                var bounds = new PowerPointLayoutBox(0, 0, 8000, 2000);
                slide.StackShapes(slide.Shapes, PowerPointShapeStackDirection.Horizontal, bounds,
                    new PowerPointShapeStackOptions {
                        SpacingEmus = 1000,
                        ClampSpacingToBounds = true
                    });

                Assert.Equal(0, a.Left);
                Assert.Equal(2750, b.Left);
                Assert.Equal(5500, c.Left);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void StackShapes_ScalesToFitBounds() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            try {
                using PowerPointPresentation presentation = PowerPointPresentation.Create(filePath);
                PowerPointSlide slide = presentation.AddSlide();
                PowerPointAutoShape a = slide.AddRectangle(0, 0, 3000, 500);
                PowerPointAutoShape b = slide.AddRectangle(0, 0, 3000, 500);

                var bounds = new PowerPointLayoutBox(0, 0, 5000, 2000);
                slide.StackShapes(slide.Shapes, PowerPointShapeStackDirection.Horizontal, bounds,
                    new PowerPointShapeStackOptions {
                        SpacingEmus = 1000,
                        ScaleToFitBounds = true
                    });

                Assert.Equal(2000, a.Width);
                Assert.Equal(2000, b.Width);
                Assert.Equal(0, a.Left);
                Assert.Equal(3000, b.Left);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }
    }
}
