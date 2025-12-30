using System;
using System.IO;
using System.Linq;
using OfficeIMO.PowerPoint;
using Xunit;

namespace OfficeIMO.Tests {
    public class PowerPointShapeAlignmentTests {
        [Fact]
        public void AlignShapes_Left_UsesSelectionBounds() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            try {
                using PowerPointPresentation presentation = PowerPointPresentation.Create(filePath);
                PowerPointSlide slide = presentation.AddSlide();
                PowerPointAutoShape a = slide.AddRectangle(1000, 0, 1000, 1000);
                PowerPointAutoShape b = slide.AddRectangle(3000, 0, 1000, 1000);
                PowerPointAutoShape c = slide.AddRectangle(5000, 0, 1000, 1000);

                slide.AlignShapes(slide.Shapes, PowerPointShapeAlignment.Left);

                Assert.All(new[] { a, b, c }, shape => Assert.Equal(1000, shape.Left));
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void AlignShapesToSlide_Center_HonorsSlideBounds() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            try {
                using PowerPointPresentation presentation = PowerPointPresentation.Create(filePath);
                PowerPointSlide slide = presentation.AddSlide();
                PowerPointAutoShape shape = slide.AddRectangle(0, 0, 1000000, 1000000);

                slide.AlignShapesToSlide(new[] { shape }, PowerPointShapeAlignment.Center);

                long expected = (presentation.SlideSize.WidthEmus - shape.Width) / 2;
                Assert.Equal(expected, shape.Left);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void AlignShapesToSlideContent_Center_RespectsMargin() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            try {
                using PowerPointPresentation presentation = PowerPointPresentation.Create(filePath);
                PowerPointSlide slide = presentation.AddSlide();
                PowerPointAutoShape shape = slide.AddRectangle(0, 0, 1000000, 1000000);
                long margin = PowerPointUnits.FromCentimeters(1);

                slide.AlignShapesToSlideContent(new[] { shape }, PowerPointShapeAlignment.Center, margin);

                PowerPointLayoutBox content = presentation.SlideSize.GetContentBox(margin);
                long expected = content.Left + (content.Width - shape.Width) / 2;
                Assert.Equal(expected, shape.Left);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void DistributeShapes_Horizontal_EvenSpacing() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            try {
                using PowerPointPresentation presentation = PowerPointPresentation.Create(filePath);
                PowerPointSlide slide = presentation.AddSlide();
                PowerPointAutoShape a = slide.AddRectangle(0, 0, 1000, 1000);
                PowerPointAutoShape b = slide.AddRectangle(4000, 0, 1000, 1000);
                PowerPointAutoShape c = slide.AddRectangle(9000, 0, 1000, 1000);

                slide.DistributeShapes(slide.Shapes, PowerPointShapeDistribution.Horizontal);

                Assert.Equal(0, a.Left);
                Assert.Equal(4500, b.Left);
                Assert.Equal(9000, c.Left);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void DistributeShapes_Horizontal_AlignsBottom() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            try {
                using PowerPointPresentation presentation = PowerPointPresentation.Create(filePath);
                PowerPointSlide slide = presentation.AddSlide();
                PowerPointAutoShape a = slide.AddRectangle(0, 0, 1000, 500);
                PowerPointAutoShape b = slide.AddRectangle(2000, 200, 1000, 700);
                PowerPointAutoShape c = slide.AddRectangle(4000, 400, 1000, 900);
                var bounds = new PowerPointLayoutBox(0, 0, 6000, 2000);

                slide.DistributeShapes(slide.Shapes, PowerPointShapeDistribution.Horizontal, bounds,
                    PowerPointShapeAlignment.Bottom);

                Assert.Equal(0, a.Left);
                Assert.Equal(2500, b.Left);
                Assert.Equal(5000, c.Left);

                Assert.Equal(bounds.Bottom - a.Height, a.Top);
                Assert.Equal(bounds.Bottom - b.Height, b.Top);
                Assert.Equal(bounds.Bottom - c.Height, c.Top);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void DistributeShapesToSlideContent_Horizontal_UsesMargin() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            try {
                using PowerPointPresentation presentation = PowerPointPresentation.Create(filePath);
                PowerPointSlide slide = presentation.AddSlide();
                PowerPointAutoShape a = slide.AddRectangle(0, 0, 1000, 1000);
                PowerPointAutoShape b = slide.AddRectangle(2000, 0, 1000, 1000);
                PowerPointAutoShape c = slide.AddRectangle(4000, 0, 1000, 1000);
                long margin = PowerPointUnits.FromCentimeters(1);

                slide.DistributeShapesToSlideContent(slide.Shapes, PowerPointShapeDistribution.Horizontal, margin);

                PowerPointLayoutBox content = presentation.SlideSize.GetContentBox(margin);
                long totalWidth = a.Width + b.Width + c.Width;
                double gap = (content.Width - totalWidth) / 2d;
                double current = content.Left;
                long expectedA = (long)Math.Round(current);
                current += a.Width + gap;
                long expectedB = (long)Math.Round(current);
                current += b.Width + gap;
                long expectedC = (long)Math.Round(current);

                Assert.Equal(expectedA, a.Left);
                Assert.Equal(expectedB, b.Left);
                Assert.Equal(expectedC, c.Left);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void DistributeShapes_Horizontal_AllowsOverlapToFitBounds() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            try {
                using PowerPointPresentation presentation = PowerPointPresentation.Create(filePath);
                PowerPointSlide slide = presentation.AddSlide();
                PowerPointAutoShape a = slide.AddRectangle(0, 0, 1000, 1000);
                PowerPointAutoShape b = slide.AddRectangle(2000, 0, 1000, 1000);
                var bounds = new PowerPointLayoutBox(0, 0, 1500, 1000);

                slide.DistributeShapes(slide.Shapes, PowerPointShapeDistribution.Horizontal, bounds);

                Assert.Equal(0, a.Left);
                Assert.Equal(500, b.Left);
                Assert.Equal(bounds.Right, b.Left + b.Width);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void DistributeShapes_Vertical_AllowsOverlapToFitBounds() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            try {
                using PowerPointPresentation presentation = PowerPointPresentation.Create(filePath);
                PowerPointSlide slide = presentation.AddSlide();
                PowerPointAutoShape a = slide.AddRectangle(0, 0, 1000, 1000);
                PowerPointAutoShape b = slide.AddRectangle(0, 2000, 1000, 1000);
                var bounds = new PowerPointLayoutBox(0, 0, 1000, 1500);

                slide.DistributeShapes(slide.Shapes, PowerPointShapeDistribution.Vertical, bounds);

                Assert.Equal(0, a.Top);
                Assert.Equal(500, b.Top);
                Assert.Equal(bounds.Bottom, b.Top + b.Height);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }
    }
}
