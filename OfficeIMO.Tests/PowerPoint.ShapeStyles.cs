using System;
using System.IO;
using System.Linq;
using OfficeIMO.PowerPoint;
using Xunit;

namespace OfficeIMO.Tests {
    public class PowerPointShapeStyles {
        [Fact]
        public void CanApplyShapeStylesAndTransforms() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");

            using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                PowerPointSlide slide = presentation.AddSlide();
                PowerPointAutoShape rect = slide.AddRectangle(1000000, 1000000, 2000000, 1000000, "Styled Rect");
                rect.FillColor = "E7F7FF";
                rect.OutlineColor = "007ACC";
                rect.OutlineWidthPoints = 2;
                rect.FillTransparency = 25;
                rect.Rotation = 15;
                rect.HorizontalFlip = true;
                rect.VerticalFlip = true;

                PowerPointAutoShape accent = slide.AddEllipse(1200000, 1200000, 1500000, 800000, "Accent");
                accent.FillColor = "FDEBD0";
                accent.SendToBack();

                presentation.Save();
            }

            using (PowerPointPresentation presentation = PowerPointPresentation.Open(filePath)) {
                PowerPointSlide slide = presentation.Slides.Single();
                var shapes = slide.Shapes.OfType<PowerPointAutoShape>().ToList();
                Assert.Equal("Accent", shapes.First().Name);

                PowerPointAutoShape rect = shapes.Single(s => s.Name == "Styled Rect");
                Assert.Equal("E7F7FF", rect.FillColor);
                Assert.Equal("007ACC", rect.OutlineColor);
                Assert.Equal(2, rect.OutlineWidthPoints ?? 0, 2);
                Assert.Equal(25, rect.FillTransparency);
                Assert.Equal(15, rect.Rotation ?? 0, 2);
                Assert.True(rect.HorizontalFlip);
                Assert.True(rect.VerticalFlip);
            }

            File.Delete(filePath);
        }
    }
}
