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
                using PowerPointPresentation presentation = PowerPointPresentation.Create(filePath);
                PowerPointSlide slide = presentation.AddSlide();
                PowerPointAutoShape a = slide.AddRectangle(0, 0, 100, 100);
                PowerPointAutoShape b = slide.AddRectangle(200, 200, 100, 100);

                slide.FitShapesToBounds(slide.Shapes, new PowerPointLayoutBox(0, 0, 600, 300));

                Assert.Equal(0, a.Left);
                Assert.Equal(0, a.Top);
                Assert.Equal(200, a.Width);
                Assert.Equal(100, a.Height);

                Assert.Equal(400, b.Left);
                Assert.Equal(200, b.Top);
                Assert.Equal(200, b.Width);
                Assert.Equal(100, b.Height);
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
                using PowerPointPresentation presentation = PowerPointPresentation.Create(filePath);
                PowerPointSlide slide = presentation.AddSlide();
                PowerPointAutoShape a = slide.AddRectangle(0, 0, 100, 100);
                PowerPointAutoShape b = slide.AddRectangle(200, 200, 100, 100);

                slide.FitShapesToBounds(slide.Shapes, new PowerPointLayoutBox(0, 0, 600, 300),
                    preserveAspect: true, center: true);

                Assert.Equal(150, a.Left);
                Assert.Equal(0, a.Top);
                Assert.Equal(100, a.Width);
                Assert.Equal(100, a.Height);

                Assert.Equal(350, b.Left);
                Assert.Equal(200, b.Top);
                Assert.Equal(100, b.Width);
                Assert.Equal(100, b.Height);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }
    }
}
