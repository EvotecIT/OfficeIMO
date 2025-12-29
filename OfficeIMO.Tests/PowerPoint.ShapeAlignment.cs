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
        public void DistributeShapes_Horizontal_EvensSpacing() {
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
    }
}
