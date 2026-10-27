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
                PowerPointAutoShape a = slide.AddRectangle(0, 0, 100, 100);
                PowerPointAutoShape b = slide.AddRectangle(200, 200, 100, 100);

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
                PowerPointAutoShape a = slide.AddRectangle(0, 0, 100, 100);
                PowerPointAutoShape b = slide.AddRectangle(200, 200, 100, 100);

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
    }
}
