using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using OfficeIMO.PowerPoint;
using Xunit;

namespace OfficeIMO.Tests {
    public class PowerPointShapeArrangeTests {
        [Fact]
        public void DuplicateShape_OffsetsAndKeepsSize() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");

            try {
                using PowerPointPresentation presentation = PowerPointPresentation.Create(filePath);
                PowerPointSlide slide = presentation.AddSlide();
                PowerPointAutoShape original = slide.AddRectangle(1000, 2000, 3000, 4000);

                PowerPointShape duplicate = slide.DuplicateShape(original, 500, 600);

                Assert.Equal(2, slide.Shapes.Count);
                Assert.Equal(original.Width, duplicate.Width);
                Assert.Equal(original.Height, duplicate.Height);
                Assert.Equal(original.Left + 500, duplicate.Left);
                Assert.Equal(original.Top + 600, duplicate.Top);
                Assert.NotEqual(original.Name, duplicate.Name);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void BringForwardAndSendBackward_ReordersShapes() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");

            try {
                using PowerPointPresentation presentation = PowerPointPresentation.Create(filePath);
                PowerPointSlide slide = presentation.AddSlide();
                PowerPointAutoShape a = slide.AddRectangle(0, 0, 1000, 1000);
                PowerPointAutoShape b = slide.AddRectangle(0, 0, 1000, 1000);
                PowerPointAutoShape c = slide.AddRectangle(0, 0, 1000, 1000);

                slide.BringForward(a);

                Assert.Equal(new PowerPointShape[] { b, a, c }, slide.Shapes.ToArray());

                slide.SendBackward(c);

                Assert.Equal(new PowerPointShape[] { b, c, a }, slide.Shapes.ToArray());
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void DuplicateGroupShape_AssignsUniqueNonVisualMetadata() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");

            try {
                using PowerPointPresentation presentation = PowerPointPresentation.Create(filePath);
                PowerPointSlide slide = presentation.AddSlide();
                slide.AddRectangle(0, 0, 1000, 1000);
                slide.AddRectangle(1500, 0, 1000, 1000);

                PowerPointGroupShape group = slide.GroupShapes(slide.Shapes);
                PowerPointGroupShape duplicate = Assert.IsType<PowerPointGroupShape>(slide.DuplicateShape(group, 500, 500));

                Assert.NotEqual(group.Name, duplicate.Name);
                presentation.Save();
            } finally {
                if (File.Exists(filePath)) {
                    using PresentationDocument document = PresentationDocument.Open(filePath, false);
                    var nonVisualProps = document.PresentationPart!.SlideParts.First().Slide
                        .Descendants<NonVisualDrawingProperties>()
                        .ToList();

                    Assert.Equal(nonVisualProps.Count, nonVisualProps.Select(property => property.Id?.Value).Distinct().Count());
                    Assert.Equal(nonVisualProps.Count, nonVisualProps.Select(property => property.Name?.Value).Distinct().Count());
                    File.Delete(filePath);
                }
            }
        }
    }
}
