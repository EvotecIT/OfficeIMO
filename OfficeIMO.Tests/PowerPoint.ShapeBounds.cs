using System;
using System.IO;
using OfficeIMO.PowerPoint;
using Xunit;

namespace OfficeIMO.Tests {
    public class PowerPointShapeBoundsTests {
        [Fact]
        public void ShapeBounds_GetSet_Works() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            try {
                using PowerPointPresentation presentation = PowerPointPresentation.Create(filePath);
                PowerPointSlide slide = presentation.AddSlide();
                PowerPointAutoShape shape = slide.AddRectangle(1000, 2000, 3000, 4000);

                PowerPointLayoutBox bounds = shape.Bounds;
                Assert.Equal(1000, bounds.Left);
                Assert.Equal(2000, bounds.Top);
                Assert.Equal(3000, bounds.Width);
                Assert.Equal(4000, bounds.Height);

                shape.Bounds = new PowerPointLayoutBox(2000, 3000, 4000, 5000);
                Assert.Equal(2000, shape.Left);
                Assert.Equal(3000, shape.Top);
                Assert.Equal(4000, shape.Width);
                Assert.Equal(5000, shape.Height);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void ShapeBounds_RightBottomSetters_MoveEdges() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            try {
                using PowerPointPresentation presentation = PowerPointPresentation.Create(filePath);
                PowerPointSlide slide = presentation.AddSlide();
                PowerPointAutoShape shape = slide.AddRectangle(1000, 1000, 500, 600);

                shape.Right = 4000;
                shape.Bottom = 5000;

                Assert.Equal(3500, shape.Left);
                Assert.Equal(4400, shape.Top);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void ShapeBounds_CenterSetters_MoveCenter() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            try {
                using PowerPointPresentation presentation = PowerPointPresentation.Create(filePath);
                PowerPointSlide slide = presentation.AddSlide();
                PowerPointAutoShape shape = slide.AddRectangle(0, 0, 1000, 2000);

                shape.CenterX = 3000;
                shape.CenterY = 4000;

                Assert.Equal(2500, shape.Left);
                Assert.Equal(3000, shape.Top);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Slide_GetShapesInBounds_SelectsExpectedShapes() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            try {
                using PowerPointPresentation presentation = PowerPointPresentation.Create(filePath);
                PowerPointSlide slide = presentation.AddSlide();
                PowerPointAutoShape a = slide.AddRectangle(0, 0, 1000, 1000);
                PowerPointAutoShape b = slide.AddRectangle(2000, 0, 1000, 1000);
                PowerPointAutoShape c = slide.AddRectangle(4000, 0, 1000, 1000);

                PowerPointLayoutBox bounds = new PowerPointLayoutBox(0, 0, 2500, 1500);
                var intersects = slide.GetShapesInBounds(bounds);
                Assert.Contains(a, intersects);
                Assert.Contains(b, intersects);
                Assert.DoesNotContain(c, intersects);

                var contained = slide.GetShapesInBounds(bounds, includePartial: false);
                Assert.Contains(a, contained);
                Assert.DoesNotContain(b, contained);
                Assert.DoesNotContain(c, contained);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }
    }
}
