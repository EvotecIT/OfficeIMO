using System;
using System.IO;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using OfficeIMO.PowerPoint;
using Xunit;

namespace OfficeIMO.Tests {
    public class PowerPointDocumentSizes {
        [Fact]
        public void SlideAndNotesSizesAreSet() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");

            using (PowerPointDocument presentation = PowerPointDocument.Create(filePath)) {
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

            using (PowerPointDocument presentation = PowerPointDocument.Create(filePath)) {
                presentation.SlideSize.Width= OfficeIMO.OfficeLength.Centimeters(25.4);
                presentation.SlideSize.Height= OfficeIMO.OfficeLength.Centimeters(14.0);
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

            using (PowerPointDocument presentation = PowerPointDocument.Create(filePath)) {
                Assert.Throws<ArgumentOutOfRangeException>(() => presentation.SlideSize.Width= OfficeIMO.OfficeLength.Emus((long)int.MaxValue + 1));
                Assert.Throws<ArgumentOutOfRangeException>(() => presentation.SlideSize.Height= OfficeIMO.OfficeLength.Emus((long)int.MinValue - 1));
            }

            File.Delete(filePath);
        }

        [Fact]
        public void CanCreateLayoutBoxes() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");

            using (PowerPointDocument presentation = PowerPointDocument.Create(filePath)) {
                PowerPointLayoutBox content = presentation.SlideSize.GetContentBox(OfficeIMO.OfficeLength.Centimeters(1.0));
                PowerPointLayoutBox[] columns = presentation.SlideSize.GetColumns(2,margin: OfficeIMO.OfficeLength.Centimeters(1.0),gutter: OfficeIMO.OfficeLength.Centimeters(1.0));
                PowerPointLayoutBox[] rows = presentation.SlideSize.GetRows(2,margin: OfficeIMO.OfficeLength.Centimeters(1.0),gutter: OfficeIMO.OfficeLength.Centimeters(1.0));

                Assert.Equal(content.Left.ToEmus(), columns[0].Left.ToEmus());
                Assert.Equal(content.Right.ToEmus(), columns[1].Right.ToEmus());
                Assert.Equal(content.Top.ToEmus(), rows[0].Top.ToEmus());
                Assert.Equal(content.Bottom.ToEmus(), rows[1].Bottom.ToEmus());
            }

            File.Delete(filePath);
        }

        [Fact]
        public void LayoutBoxCanSplitRowsAndColumns() {
            PowerPointLayoutBox box = new OfficeIMO.PowerPoint.PowerPointLayoutBox(OfficeIMO.OfficeLength.Centimeters(1.0),OfficeIMO.OfficeLength.Centimeters(2.0),OfficeIMO.OfficeLength.Centimeters(20.0),OfficeIMO.OfficeLength.Centimeters(10.0));
            PowerPointLayoutBox[] columns = box.SplitColumns(2,gutter: OfficeIMO.OfficeLength.Centimeters(1.0));
            PowerPointLayoutBox[] rows = box.SplitRows(2,gutter: OfficeIMO.OfficeLength.Centimeters(1.0));

            Assert.Equal(box.Left.ToEmus(), columns[0].Left.ToEmus());
            Assert.Equal(box.Right.ToEmus(), columns[1].Right.ToEmus());
            Assert.Equal(box.Top.ToEmus(), rows[0].Top.ToEmus());
            Assert.Equal(box.Bottom.ToEmus(), rows[1].Bottom.ToEmus());
        }

        [Fact]
        public void CanSetSlideSizePreset() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");

            using (PowerPointDocument presentation = PowerPointDocument.Create(filePath)) {
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

        [Fact]
        public void Save_RefreshesPresentationMetadata() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");

            try {
                using (PowerPointDocument presentation = PowerPointDocument.Create(filePath)) {
                    presentation.SlideSize.SetPreset(PowerPointSlideSizePreset.Screen4x3);

                    PowerPointSlide slide1 = presentation.AddSlide();
                    slide1.Notes.Text = "Notes";

                    presentation.AddSlide();
                    presentation.AddSlide().Hide();

                    presentation.Save();
                }

                using (PresentationDocument document = PresentationDocument.Open(filePath, false)) {
                    var properties = document.ExtendedFilePropertiesPart!.Properties!;

                    Assert.Equal("3", properties.Slides!.Text);
                    Assert.Equal("1", properties.Notes!.Text);
                    Assert.Equal("1", properties.HiddenSlides!.Text);
                    Assert.Equal("Standard", properties.PresentationFormat!.Text);
                    Assert.NotNull(document.PackageProperties.Modified);
                }
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }
    }
}
