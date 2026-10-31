using System;
using System.IO;
using OfficeIMO.PowerPoint;
using Xunit;

namespace OfficeIMO.Tests {
    public class PowerPointShapeFitTests {
        [Fact]
        public void FitShapesToBounds_ScalesNonUniform() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            try {
                using PowerPointDocument presentation = PowerPointDocument.Create(filePath);
                PowerPointSlide slide = presentation.AddSlide();
                PowerPointAutoShape a = slide.AddRectangle(OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(100), OfficeIMO.OfficeLength.Emus(100));
                PowerPointAutoShape b = slide.AddRectangle(OfficeIMO.OfficeLength.Emus(200), OfficeIMO.OfficeLength.Emus(200), OfficeIMO.OfficeLength.Emus(100), OfficeIMO.OfficeLength.Emus(100));

                slide.FitShapesToBounds(slide.Shapes, new PowerPointLayoutBox(OfficeIMO.OfficeLength.Emus(0),OfficeIMO.OfficeLength.Emus(0),OfficeIMO.OfficeLength.Emus(600),OfficeIMO.OfficeLength.Emus(300)));

                Assert.Equal(0, a.Left.ToEmus());
                Assert.Equal(0, a.Top.ToEmus());
                Assert.Equal(200, a.Width.ToEmus());
                Assert.Equal(100, a.Height.ToEmus());

                Assert.Equal(400, b.Left.ToEmus());
                Assert.Equal(200, b.Top.ToEmus());
                Assert.Equal(200, b.Width.ToEmus());
                Assert.Equal(100, b.Height.ToEmus());
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void FitShapesToBounds_PreservesAspectAndCenters() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            try {
                using PowerPointDocument presentation = PowerPointDocument.Create(filePath);
                PowerPointSlide slide = presentation.AddSlide();
                PowerPointAutoShape a = slide.AddRectangle(OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(100), OfficeIMO.OfficeLength.Emus(100));
                PowerPointAutoShape b = slide.AddRectangle(OfficeIMO.OfficeLength.Emus(200), OfficeIMO.OfficeLength.Emus(200), OfficeIMO.OfficeLength.Emus(100), OfficeIMO.OfficeLength.Emus(100));

                slide.FitShapesToBounds(slide.Shapes, new PowerPointLayoutBox(OfficeIMO.OfficeLength.Emus(0),OfficeIMO.OfficeLength.Emus(0),OfficeIMO.OfficeLength.Emus(600),OfficeIMO.OfficeLength.Emus(300)),
                    preserveAspect: true, center: true);

                Assert.Equal(150, a.Left.ToEmus());
                Assert.Equal(0, a.Top.ToEmus());
                Assert.Equal(100, a.Width.ToEmus());
                Assert.Equal(100, a.Height.ToEmus());

                Assert.Equal(350, b.Left.ToEmus());
                Assert.Equal(200, b.Top.ToEmus());
                Assert.Equal(100, b.Width.ToEmus());
                Assert.Equal(100, b.Height.ToEmus());
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }
        [Fact]
        public void FitShapesToBounds_PreservesSignedPositionsAndMidpointToEvenRounding() {
            using PowerPointDocument presentation = PowerPointDocument.Create();
            PowerPointSlide slide = presentation.AddSlide();
            PowerPointAutoShape first = slide.AddRectangle(OfficeIMO.OfficeLength.Zero, OfficeIMO.OfficeLength.Zero,
                OfficeIMO.OfficeLength.Emus(1), OfficeIMO.OfficeLength.Emus(1));
            PowerPointAutoShape second = slide.AddRectangle(OfficeIMO.OfficeLength.Emus(1), OfficeIMO.OfficeLength.Zero,
                OfficeIMO.OfficeLength.Emus(1), OfficeIMO.OfficeLength.Emus(1));

            slide.FitShapesToBounds(slide.Shapes, new PowerPointLayoutBox(OfficeIMO.OfficeLength.Emus(-1),
                OfficeIMO.OfficeLength.Emus(-3), OfficeIMO.OfficeLength.Emus(3), OfficeIMO.OfficeLength.Emus(1)));

            Assert.Equal(-1, first.Left.ToEmus());
            Assert.Equal(0, second.Left.ToEmus());
            Assert.Equal(-3, first.Top.ToEmus());
            Assert.Equal(2, first.Width.ToEmus());
            Assert.Equal(2, second.Width.ToEmus());
        }

        [Fact]
        public void FitShapesToSlideContent_UsesPhysicalMarginAndRejectsNegativeMarginBeforeMutation() {
            using PowerPointDocument presentation = PowerPointDocument.Create();
            PowerPointSlide slide = presentation.AddSlide();
            PowerPointAutoShape shape = slide.AddRectangle(OfficeIMO.OfficeLength.Zero, OfficeIMO.OfficeLength.Zero,
                OfficeIMO.OfficeLength.Points(10), OfficeIMO.OfficeLength.Points(20));
            OfficeIMO.OfficeLength margin = OfficeIMO.OfficeLength.Centimeters(1);
            PowerPointLayoutBox expected = presentation.SlideSize.GetContentBox(margin);

            slide.FitShapesToSlideContent(slide.Shapes, margin);
            Assert.Equal(expected.Left, shape.Left);
            Assert.Equal(expected.Top, shape.Top);
            Assert.Equal(expected.Width, shape.Width);
            Assert.Equal(expected.Height, shape.Height);

            Assert.Throws<ArgumentOutOfRangeException>(() => slide.FitShapesToSlideContent(slide.Shapes,
                OfficeIMO.OfficeLength.Emus(-1)));
            Assert.Equal(expected.Left, shape.Left);
            Assert.Equal(expected.Width, shape.Width);

            var beforeSave = slide.Shapes.Select(item => (item.Left, item.Top, item.Width, item.Height)).ToArray();
            using var saved = new MemoryStream(presentation.ToBytes());
            using PowerPointDocument loaded = PowerPointDocument.Load(saved);
            Assert.Equal(beforeSave, loaded.Slides[0].Shapes.Select(item => (item.Left, item.Top, item.Width, item.Height)).ToArray());
        }

    }
}
