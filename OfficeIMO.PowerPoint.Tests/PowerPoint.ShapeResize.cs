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
                using PowerPointDocument presentation = PowerPointDocument.Create(filePath);
                PowerPointSlide slide = presentation.AddSlide();
                PowerPointAutoShape a = slide.AddRectangle(OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(1000), OfficeIMO.OfficeLength.Emus(2000));
                PowerPointAutoShape b = slide.AddRectangle(OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(2000), OfficeIMO.OfficeLength.Emus(3000));
                PowerPointAutoShape c = slide.AddRectangle(OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(1500), OfficeIMO.OfficeLength.Emus(4000));

                slide.ResizeShapes(slide.Shapes, PowerPointShapeSizeDimension.Width, PowerPointShapeSizeReference.Largest);

                Assert.Equal(2000, a.Width.ToEmus());
                Assert.Equal(2000, b.Width.ToEmus());
                Assert.Equal(2000, c.Width.ToEmus());
                Assert.Equal(2000, a.Height.ToEmus());
                Assert.Equal(3000, b.Height.ToEmus());
                Assert.Equal(4000, c.Height.ToEmus());
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
                using PowerPointDocument presentation = PowerPointDocument.Create(filePath);
                PowerPointSlide slide = presentation.AddSlide();
                PowerPointAutoShape a = slide.AddRectangle(OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(1000), OfficeIMO.OfficeLength.Emus(2000));
                PowerPointAutoShape b = slide.AddRectangle(OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(3000), OfficeIMO.OfficeLength.Emus(4000));

                slide.ResizeShapes(slide.Shapes, PowerPointShapeSizeDimension.Both, PowerPointShapeSizeReference.Average);

                Assert.Equal(2000, a.Width.ToEmus());
                Assert.Equal(2000, b.Width.ToEmus());
                Assert.Equal(3000, a.Height.ToEmus());
                Assert.Equal(3000, b.Height.ToEmus());
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
                using PowerPointDocument presentation = PowerPointDocument.Create(filePath);
                PowerPointSlide slide = presentation.AddSlide();
                PowerPointAutoShape a = slide.AddRectangle(OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(1000), OfficeIMO.OfficeLength.Emus(500));
                PowerPointAutoShape b = slide.AddRectangle(OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(1000), OfficeIMO.OfficeLength.Emus(900));

                slide.ResizeShapes(slide.Shapes, PowerPointShapeSizeDimension.Height, PowerPointShapeSizeReference.First);

                Assert.Equal(500, a.Height.ToEmus());
                Assert.Equal(500, b.Height.ToEmus());
                Assert.Equal(1000, a.Width.ToEmus());
                Assert.Equal(1000, b.Width.ToEmus());
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }
        [Fact]
        public void ResizeShapes_NullDimensionsRetainSizeAndZeroIsExplicit() {
            using PowerPointDocument presentation = PowerPointDocument.Create();
            PowerPointSlide slide = presentation.AddSlide();
            PowerPointAutoShape shape = slide.AddRectangle(OfficeIMO.OfficeLength.Zero, OfficeIMO.OfficeLength.Zero,
                OfficeIMO.OfficeLength.Points(10), OfficeIMO.OfficeLength.Points(20));

            slide.ResizeShapes(slide.Shapes, OfficeIMO.OfficeLength.Inches(1), null);
            Assert.Equal(OfficeIMO.OfficeLength.Inches(1), shape.Width);
            Assert.Equal(OfficeIMO.OfficeLength.Points(20), shape.Height);

            slide.ResizeShapes(slide.Shapes, null, null);
            Assert.Equal(OfficeIMO.OfficeLength.Inches(1), shape.Width);
            Assert.Equal(OfficeIMO.OfficeLength.Points(20), shape.Height);

            slide.ResizeShapes(slide.Shapes, null, OfficeIMO.OfficeLength.Zero);
            Assert.Equal(OfficeIMO.OfficeLength.Inches(1), shape.Width);
            Assert.Equal(OfficeIMO.OfficeLength.Zero, shape.Height);
        }

        [Fact]
        public void ResizeShapes_RejectsBothNegativeDimensionsBeforeChangingAnyShape() {
            using PowerPointDocument presentation = PowerPointDocument.Create();
            PowerPointSlide slide = presentation.AddSlide();
            PowerPointAutoShape shape = slide.AddRectangle(OfficeIMO.OfficeLength.Zero, OfficeIMO.OfficeLength.Zero,
                OfficeIMO.OfficeLength.Points(10), OfficeIMO.OfficeLength.Points(20));

            Assert.Throws<ArgumentOutOfRangeException>(() => slide.ResizeShapes(slide.Shapes,
                OfficeIMO.OfficeLength.Points(30), OfficeIMO.OfficeLength.Emus(-1)));
            Assert.Throws<ArgumentOutOfRangeException>(() => slide.ResizeShapes(slide.Shapes,
                OfficeIMO.OfficeLength.Emus(-1), OfficeIMO.OfficeLength.Points(30)));
            Assert.Equal(OfficeIMO.OfficeLength.Points(10), shape.Width);
            Assert.Equal(OfficeIMO.OfficeLength.Points(20), shape.Height);
        }

    }
}
