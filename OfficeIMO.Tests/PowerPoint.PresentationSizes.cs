using System;
using System.IO;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.PowerPoint;
using Xunit;

namespace OfficeIMO.Tests {
    public class PowerPointPresentationSizes {
        [Fact]
        public void SlideAndNotesSizesAreSet() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");

            using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                presentation.AddSlide();
                presentation.Save();
            }

            using (PresentationDocument document = PresentationDocument.Open(filePath, false)) {
                Assert.NotNull(document.PresentationPart?.Presentation?.SlideSize);
                Assert.NotNull(document.PresentationPart?.Presentation?.NotesSize);
            }

            File.Delete(filePath);
        }

        [Fact]
        public void CanSetSlideSizeUsingCentimeters() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");

            using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                presentation.SlideSize.WidthCm = 25.4;
                presentation.SlideSize.HeightCm = 14.0;
                presentation.Save();
            }

            using (PresentationDocument document = PresentationDocument.Open(filePath, false)) {
                var slideSize = document.PresentationPart?.Presentation?.SlideSize;
                Assert.NotNull(slideSize);
                Assert.Equal((int)PowerPointUnits.FromCentimeters(25.4), slideSize!.Cx!.Value);
                Assert.Equal((int)PowerPointUnits.FromCentimeters(14.0), slideSize!.Cy!.Value);
            }

            File.Delete(filePath);
        }

        [Fact]
        public void CanCreateLayoutBoxes() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");

            using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                PowerPointLayoutBox content = presentation.SlideSize.GetContentBoxCm(1.0);
                PowerPointLayoutBox[] columns = presentation.SlideSize.GetColumnsCm(2, marginCm: 1.0, gutterCm: 1.0);
                PowerPointLayoutBox[] rows = presentation.SlideSize.GetRowsCm(2, marginCm: 1.0, gutterCm: 1.0);

                Assert.Equal(content.Left, columns[0].Left);
                Assert.Equal(content.Right, columns[1].Right);
                Assert.Equal(content.Top, rows[0].Top);
                Assert.Equal(content.Bottom, rows[1].Bottom);
            }

            File.Delete(filePath);
        }
    }
}
