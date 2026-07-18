using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Validation;
using OfficeIMO.PowerPoint;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.Tests {
    public class PowerPointSlidesManagement {
        [Fact]
        public void CanRemoveSlide() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");

            using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                presentation.AddSlide().AddTextBox("Slide 1");
                presentation.AddSlide().AddTextBox("Slide 2");
                presentation.AddSlide().AddTextBox("Slide 3");

                presentation.RemoveSlide(1);

                Assert.Equal(2, presentation.Slides.Count);
                Assert.Equal("Slide 1", presentation.Slides[0].TextBoxes.First().Text);
                Assert.Equal("Slide 3", presentation.Slides[1].TextBoxes.First().Text);

                presentation.Save();
            }

            using (PowerPointPresentation presentation = PowerPointPresentation.Load(filePath)) {
                Assert.Equal(2, presentation.Slides.Count);
                Assert.Equal("Slide 1", presentation.Slides[0].TextBoxes.First().Text);
                Assert.Equal("Slide 3", presentation.Slides[1].TextBoxes.First().Text);
            }

            File.Delete(filePath);
        }

        [Fact]
        public void RemoveSlide_CleansInboundInternalHyperlinks() {
            using PowerPointPresentation presentation =
                PowerPointPresentation.Create();
            PowerPointSlide source = presentation.AddSlide();
            PowerPointTextRun run = source.AddTextBox("Open target")
                .Paragraphs.Single().Runs.Single();
            PowerPointSlide target = presentation.AddSlide();
            target.AddTextBox("Removed target content");
            run.SetHyperlink(target, "Target slide");
            Assert.Contains(source.SlidePart.Parts, pair =>
                ReferenceEquals(pair.OpenXmlPart, target.SlidePart));

            presentation.RemoveSlide(1);

            Assert.Single(presentation.Slides);
            Assert.DoesNotContain(source.SlidePart.Parts, pair =>
                ReferenceEquals(pair.OpenXmlPart, target.SlidePart));
            Assert.Empty(source.SlidePart.Slide
                .Descendants<A.HyperlinkOnClick>());
            Assert.Empty(presentation.ValidateDocument());
            byte[] bytes = presentation.ToBytes();
            using var stream = new MemoryStream(bytes, writable: false);
            using PresentationDocument document =
                PresentationDocument.Open(stream, false);
            SlidePart remaining = Assert.Single(document.PresentationPart!
                .SlideParts);
            Assert.DoesNotContain(remaining.Parts,
                pair => pair.OpenXmlPart is SlidePart);
            Assert.Empty(remaining.Slide
                .Descendants<A.HyperlinkOnClick>());
        }

        [Fact]
        public void RemovingSlidesDownToOneKeepsPresentationValid() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");

            using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                presentation.AddSlide().AddTextBox("Slide 1");
                presentation.AddSlide().AddTextBox("Slide 2");
                presentation.AddSlide().AddTextBox("Slide 3");

                presentation.RemoveSlide(2);
                presentation.RemoveSlide(1);

                Assert.Single(presentation.Slides);
                Assert.True(presentation.DocumentIsValid);

                presentation.Save();
            }

            using (PowerPointPresentation presentation = PowerPointPresentation.Load(filePath)) {
                Assert.Single(presentation.Slides);
                Assert.True(presentation.DocumentIsValid);
            }

            File.Delete(filePath);
        }

        [Fact]
        public void CanMoveSlide() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");

            using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                presentation.AddSlide().AddTextBox("Slide 1");
                presentation.AddSlide().AddTextBox("Slide 2");
                presentation.AddSlide().AddTextBox("Slide 3");

                presentation.MoveSlide(0, 2);

                Assert.Equal("Slide 2", presentation.Slides[0].TextBoxes.First().Text);
                Assert.Equal("Slide 3", presentation.Slides[1].TextBoxes.First().Text);
                Assert.Equal("Slide 1", presentation.Slides[2].TextBoxes.First().Text);

                presentation.Save();
            }

            using (PowerPointPresentation presentation = PowerPointPresentation.Load(filePath)) {
                Assert.Equal(3, presentation.Slides.Count);
                Assert.Equal("Slide 2", presentation.Slides[0].TextBoxes.First().Text);
                Assert.Equal("Slide 3", presentation.Slides[1].TextBoxes.First().Text);
                Assert.Equal("Slide 1", presentation.Slides[2].TextBoxes.First().Text);
            }

            File.Delete(filePath);
        }

        [Fact]
        public void RemovingInvalidSlideThrows() {
            using PowerPointPresentation presentation = PowerPointPresentation.Create(Path.GetTempFileName());
            Assert.Throws<ArgumentOutOfRangeException>(() => presentation.RemoveSlide(0));
        }

        [Fact]
        public void RemovingLastSlideLeavesAValidEmptyPresentation() {
            using PowerPointPresentation presentation = PowerPointPresentation.Create(Path.GetTempFileName());
            presentation.AddSlide();

            presentation.RemoveSlide(0);

            Assert.Empty(presentation.Slides);
            Assert.Empty(presentation.ValidateDocument());
        }

        [Fact]
        public void MovingInvalidSlideThrows() {
            using PowerPointPresentation presentation = PowerPointPresentation.Create(Path.GetTempFileName());
            presentation.AddSlide();
            Assert.Throws<ArgumentOutOfRangeException>(() => presentation.MoveSlide(0, 1));
            Assert.Throws<ArgumentOutOfRangeException>(() => presentation.MoveSlide(1, 0));
        }

        [Fact]
        public void CanDuplicateSlide() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            string imagePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Images", "BackgroundImage.png");

            using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                PowerPointSlide source = presentation.AddSlide();
                source.AddTextBox("Source slide");
                source.AddPicture(imagePath);
                source.Notes.Text = "Speaker notes";
                source.Hidden = true;

                PowerPointSlide duplicate = presentation.DuplicateSlide(0);

                Assert.Equal(2, presentation.Slides.Count);
                Assert.Equal("Source slide", duplicate.TextBoxes.First().Text);
                Assert.Single(duplicate.Pictures);
                Assert.True(duplicate.Hidden);
                Assert.Equal("Speaker notes", duplicate.Notes.Text);

                presentation.Save();
                Assert.Empty(presentation.ValidateDocument());
            }

            using (PowerPointPresentation presentation = PowerPointPresentation.Load(filePath)) {
                Assert.Equal(2, presentation.Slides.Count);
                Assert.True(presentation.Slides[1].Hidden);
                Assert.Equal("Source slide", presentation.Slides[1].TextBoxes.First().Text);
                Assert.Equal("Speaker notes", presentation.Slides[1].Notes.Text);
            }

            File.Delete(filePath);
        }

        [Fact]
        public void CanDuplicateSlideWithChart() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");

            using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                PowerPointSlide source = presentation.AddSlide();
                source.AddChart();

                presentation.DuplicateSlide(0);
                Assert.Equal(2, presentation.Slides.Count);
                Assert.Single(presentation.Slides[0].Charts);
                Assert.Single(presentation.Slides[1].Charts);

                presentation.Save();
            }

            using (PowerPointPresentation presentation = PowerPointPresentation.Load(filePath)) {
                Assert.Equal(2, presentation.Slides.Count);
                Assert.Single(presentation.Slides[0].Charts);
                Assert.Single(presentation.Slides[1].Charts);
            }

            File.Delete(filePath);
        }

        [Fact]
        public void CanImportSlideFromAnotherPresentation() {
            string sourcePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            string targetPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            string imagePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Images", "BackgroundImage.png");

            using (PowerPointPresentation source = PowerPointPresentation.Create(sourcePath)) {
                PowerPointSlide sourceSlide = source.AddSlide();
                sourceSlide.AddTextBox("Imported slide");
                sourceSlide.AddPicture(imagePath);
                sourceSlide.Notes.Text = "Imported notes";
                sourceSlide.Hidden = true;
                source.Save();
                Assert.Empty(source.ValidateDocument());

                using (PowerPointPresentation target = PowerPointPresentation.Create(targetPath)) {
                    PowerPointSlide imported = target.ImportSlide(source, 0);

                    Assert.Single(target.Slides);
                    Assert.Equal("Imported slide", imported.TextBoxes.First().Text);
                    Assert.Single(imported.Pictures);
                    Assert.True(imported.Hidden);
                    Assert.Equal("Imported notes", imported.Notes.Text);

                    target.Save();
                    Assert.Empty(target.ValidateDocument());
                }
            }

            using (PowerPointPresentation target = PowerPointPresentation.Load(targetPath)) {
                Assert.Single(target.Slides);
                Assert.Equal("Imported slide", target.Slides[0].TextBoxes.First().Text);
                Assert.Single(target.Slides[0].Pictures);
                Assert.True(target.Slides[0].Hidden);
                Assert.Equal("Imported notes", target.Slides[0].Notes.Text);
            }

            File.Delete(sourcePath);
            File.Delete(targetPath);
        }

        [Fact]
        public void HiddenSlideUsesSlideShowAttributeAndValidates() {
            string filePath = CreateTempFilePath(".pptx");

            try {
                using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                    presentation.AddSlide().AddTextBox("Visible slide");
                    PowerPointSlide hiddenSlide = presentation.AddSlide();
                    hiddenSlide.AddTextBox("Hidden slide");
                    hiddenSlide.Hide();

                    presentation.Save();
                    Assert.Empty(presentation.ValidateDocument());
                }

                using (PresentationDocument document = PresentationDocument.Open(filePath, false)) {
                    OpenXmlValidator validator = new();
                    Assert.Empty(validator.Validate(document));

                    PresentationPart presentationPart = document.PresentationPart!;
                    SlideId hiddenSlideId = presentationPart.Presentation!.SlideIdList!.Elements<SlideId>().Last();
                    Assert.DoesNotContain(hiddenSlideId.GetAttributes(),
                        attribute => attribute.LocalName == "show" && string.IsNullOrEmpty(attribute.NamespaceUri));

                    SlidePart hiddenSlidePart = (SlidePart)presentationPart.GetPartById(hiddenSlideId.RelationshipId!);
                    Assert.False(hiddenSlidePart.Slide!.Show!.Value);
                }

                using (PowerPointPresentation presentation = PowerPointPresentation.Load(filePath)) {
                    Assert.False(presentation.Slides[0].Hidden);
                    Assert.True(presentation.Slides[1].Hidden);
                }
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void ImportSlide_PreservesRichNotesMarkup() {
            string sourcePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            string targetPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");

            try {
                using (PowerPointPresentation source = PowerPointPresentation.Create(sourcePath)) {
                    PowerPointSlide sourceSlide = source.AddSlide();
                    sourceSlide.AddTextBox("Imported slide");
                    sourceSlide.Notes.Text = "Imported notes";
                    source.Save();
                }

                using (PresentationDocument document = PresentationDocument.Open(sourcePath, true)) {
                    NotesSlidePart notesPart = document.PresentationPart!.SlideParts.Single().NotesSlidePart!;
                    ShapeTree shapeTree = notesPart.NotesSlide!.CommonSlideData!.ShapeTree!;
                    shapeTree.Append(new Shape(
                        new NonVisualShapeProperties(
                            new NonVisualDrawingProperties { Id = 99U, Name = "Custom Notes Shape" },
                            new NonVisualShapeDrawingProperties(),
                            new ApplicationNonVisualDrawingProperties()),
                        new ShapeProperties(),
                        new TextBody(
                            new A.BodyProperties(),
                            new A.ListStyle(),
                            new A.Paragraph(new A.Run(new A.Text("Extra notes shape"))))));
                    notesPart.NotesSlide.Save();
                }

                using (PowerPointPresentation source = PowerPointPresentation.Load(sourcePath))
                using (PowerPointPresentation target = PowerPointPresentation.Create(targetPath)) {
                    PowerPointSlide imported = target.ImportSlide(source, 0);

                    Assert.Equal("Imported notes", imported.Notes.Text);
                    target.Save();
                    Assert.Empty(target.ValidateDocument());
                }

                using (PresentationDocument document = PresentationDocument.Open(targetPath, false)) {
                    OpenXmlValidator validator = new();
                    Assert.Empty(validator.Validate(document));

                    NotesSlidePart notesPart = document.PresentationPart!.SlideParts.Single().NotesSlidePart!;
                    Shape[] shapes = notesPart.NotesSlide!.CommonSlideData!.ShapeTree!.Elements<Shape>().ToArray();
                    Assert.Contains(shapes, shape =>
                        shape.NonVisualShapeProperties?.NonVisualDrawingProperties?.Name?.Value == "Custom Notes Shape");
                    Assert.Contains("Extra notes shape", notesPart.NotesSlide.OuterXml, StringComparison.Ordinal);
                }
            } finally {
                if (File.Exists(sourcePath)) {
                    File.Delete(sourcePath);
                }

                if (File.Exists(targetPath)) {
                    File.Delete(targetPath);
                }
            }
        }

        private static string CreateTempFilePath(string extension) {
            string path = Path.GetTempFileName();
            File.Delete(path);
            return Path.ChangeExtension(path, extension);
        }
    }
}
