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
                using PowerPointDocument presentation = PowerPointDocument.Create(filePath);
                PowerPointSlide slide = presentation.AddSlide();
                PowerPointAutoShape shape = slide.AddRectangle(1000, 2000, 3000, 4000);

                PowerPointLayoutBox bounds = shape.Bounds;
                Assert.Equal(1000, bounds.Left.ToEmus());
                Assert.Equal(2000, bounds.Top.ToEmus());
                Assert.Equal(3000, bounds.Width.ToEmus());
                Assert.Equal(4000, bounds.Height.ToEmus());

                shape.Bounds = new PowerPointLayoutBox(OfficeIMO.OfficeLength.Emus(2000),OfficeIMO.OfficeLength.Emus(3000),OfficeIMO.OfficeLength.Emus(4000),OfficeIMO.OfficeLength.Emus(5000));
                Assert.Equal(2000, shape.Left.ToEmus());
                Assert.Equal(3000, shape.Top.ToEmus());
                Assert.Equal(4000, shape.Width.ToEmus());
                Assert.Equal(5000, shape.Height.ToEmus());
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
                using PowerPointDocument presentation = PowerPointDocument.Create(filePath);
                PowerPointSlide slide = presentation.AddSlide();
                PowerPointAutoShape shape = slide.AddRectangle(1000, 1000, 500, 600);

                shape.Right= OfficeIMO.OfficeLength.Emus(4000);
                shape.Bottom= OfficeIMO.OfficeLength.Emus(5000);

                Assert.Equal(3500, shape.Left.ToEmus());
                Assert.Equal(4400, shape.Top.ToEmus());
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
                using PowerPointDocument presentation = PowerPointDocument.Create(filePath);
                PowerPointSlide slide = presentation.AddSlide();
                PowerPointAutoShape shape = slide.AddRectangle(0, 0, 1000, 2000);

                shape.CenterX= OfficeIMO.OfficeLength.Emus(3000);
                shape.CenterY= OfficeIMO.OfficeLength.Emus(4000);

                Assert.Equal(2500, shape.Left.ToEmus());
                Assert.Equal(3000, shape.Top.ToEmus());
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
                using PowerPointDocument presentation = PowerPointDocument.Create(filePath);
                PowerPointSlide slide = presentation.AddSlide();
                PowerPointAutoShape a = slide.AddRectangle(0, 0, 1000, 1000);
                PowerPointAutoShape b = slide.AddRectangle(2000, 0, 1000, 1000);
                PowerPointAutoShape c = slide.AddRectangle(4000, 0, 1000, 1000);

                PowerPointLayoutBox bounds = new PowerPointLayoutBox(OfficeIMO.OfficeLength.Emus(0),OfficeIMO.OfficeLength.Emus(0),OfficeIMO.OfficeLength.Emus(2500),OfficeIMO.OfficeLength.Emus(1500));
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
