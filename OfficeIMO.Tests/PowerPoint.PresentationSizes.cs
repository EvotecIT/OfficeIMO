using System;
using System.IO;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
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
        public void SlideSizeRejectsValuesOutsideIntRange() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");

            using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                Assert.Throws<ArgumentOutOfRangeException>(() => presentation.SlideSize.WidthEmus = (long)int.MaxValue + 1);
                Assert.Throws<ArgumentOutOfRangeException>(() => presentation.SlideSize.HeightEmus = (long)int.MinValue - 1);
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

        [Fact]
        public void LayoutBoxCanSplitRowsAndColumns() {
            PowerPointLayoutBox box = PowerPointLayoutBox.FromCentimeters(1.0, 2.0, 20.0, 10.0);
            PowerPointLayoutBox[] columns = box.SplitColumnsCm(2, gutterCm: 1.0);
            PowerPointLayoutBox[] rows = box.SplitRowsCm(2, gutterCm: 1.0);

            Assert.Equal(box.Left, columns[0].Left);
            Assert.Equal(box.Right, columns[1].Right);
            Assert.Equal(box.Top, rows[0].Top);
            Assert.Equal(box.Bottom, rows[1].Bottom);
        }

        [Fact]
        public void CanSetSlideSizePreset() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");

            using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                presentation.SlideSize.SetPreset(PowerPointSlideSizePreset.Screen4x3);
                presentation.Save();
            }

            using (PresentationDocument document = PresentationDocument.Open(filePath, false)) {
                SlideSize slideSize = document.PresentationPart!.Presentation!.SlideSize!;
                Assert.Equal(SlideSizeValues.Screen4x3, slideSize.Type!.Value);
                Assert.Equal((int)PowerPointUnits.FromInches(10), slideSize.Cx!.Value);
                Assert.Equal((int)PowerPointUnits.FromInches(7.5), slideSize.Cy!.Value);
            }

            File.Delete(filePath);
        }
    }
}
