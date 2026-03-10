using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using OfficeIMO.PowerPoint;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.Tests {
    public class PowerPointAdvancedFeatures {
        [Fact]
        public void CanHandleBackgroundFormattingTransitionsAndCharts() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            string imagePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Images", "BackgroundImage.png");

            using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                PowerPointSlide slide = presentation.AddSlide();
                PowerPointTextBox text = slide.AddTextBox("Test");
                slide.AddPicture(imagePath);
                slide.AddTable(2, 2);
                slide.AddChart();
                slide.Notes.Text = "Notes";

                slide.BackgroundColor = "FF0000";
                text.FillColor = "00FF00";
                slide.Transition = SlideTransition.Fade;

                presentation.Save();
            }

            using (PowerPointPresentation presentation = PowerPointPresentation.Open(filePath)) {
                PowerPointSlide slide = presentation.Slides.Single();
                Assert.Equal("FF0000", slide.BackgroundColor);
                Assert.Equal(SlideTransition.Fade, slide.Transition);
                Assert.Single(slide.TextBoxes);
                Assert.Single(slide.Pictures);
                Assert.Single(slide.Tables);
                Assert.Single(slide.Charts);
                Assert.Equal("00FF00", slide.TextBoxes.First().FillColor);
                Assert.Equal("Notes", slide.Notes.Text);
            }

            File.Delete(filePath);
        }

        [Fact]
        public void CanAddMultipleChartsWithUniqueAxisIds() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");

            try {
                using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                    PowerPointSlide slide = presentation.AddSlide();
                    for (int i = 0; i < 3; i++) {
                        slide.AddChart();
                    }

                    presentation.Save();
                    Assert.Empty(presentation.ValidateDocument());
                }

                using (PowerPointPresentation presentation = PowerPointPresentation.Open(filePath)) {
                    PowerPointSlide slide = presentation.Slides.Single();
                    Assert.Equal(3, slide.Charts.Count());
                    Assert.Empty(presentation.ValidateDocument());
                }

                using (PresentationDocument presentationDocument = PresentationDocument.Open(filePath, false)) {
                    PresentationPart? presentationPart = presentationDocument.PresentationPart;
                    Assert.NotNull(presentationPart);

                    HashSet<uint> axisIds = new();
                    foreach (ChartPart chartPart in presentationPart!.SlideParts.SelectMany(s => s.ChartParts)) {
                        Chart? chart = chartPart.ChartSpace?.GetFirstChild<Chart>();
                        Assert.NotNull(chart);

                        IEnumerable<uint> axisValues = (chart!.PlotArea?.Elements<OpenXmlCompositeElement>()
                            ?? Enumerable.Empty<OpenXmlCompositeElement>())
                            .Where(element => element is CategoryAxis || element is ValueAxis || element is SeriesAxis || element is DateAxis)
                            .SelectMany(element => element.Elements<AxisId>())
                            .Select(axis => axis.Val?.Value)
                            .OfType<uint>();

                        foreach (uint axisId in axisValues) {
                            Assert.True(axisIds.Add(axisId), $"Duplicate axis id {axisId} found.");
                        }
                    }

                    Assert.Equal(6, axisIds.Count);
                }
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void BackgroundColor_ReplacesExistingBackgroundImageFill() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            string imagePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Images", "BackgroundImage.png");

            try {
                using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                    PowerPointSlide slide = presentation.AddSlide();
                    slide.SetBackgroundImage(imagePath);
                    slide.BackgroundColor = "112233";

                    presentation.Save();
                    Assert.Empty(presentation.ValidateDocument());
                }

                using (PowerPointPresentation presentation = PowerPointPresentation.Open(filePath)) {
                    PowerPointSlide slide = presentation.Slides.Single();
                    Assert.Equal("112233", slide.BackgroundColor);
                    Assert.Empty(presentation.ValidateDocument());
                }

                using (PresentationDocument document = PresentationDocument.Open(filePath, false)) {
                    SlidePart slidePart = document.PresentationPart!.SlideParts.First();
                    DocumentFormat.OpenXml.Presentation.BackgroundProperties properties =
                        slidePart.Slide.CommonSlideData!.Background!.BackgroundProperties!;

                    Assert.Null(properties.GetFirstChild<A.BlipFill>());
                    Assert.Equal("112233", properties.GetFirstChild<A.SolidFill>()?.RgbColorModelHex?.Val?.Value);
                }
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void ClearingBackgroundColor_DoesNotRemoveBackgroundImage() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            string imagePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Images", "BackgroundImage.png");

            try {
                using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                    PowerPointSlide slide = presentation.AddSlide();
                    slide.SetBackgroundImage(imagePath);
                    slide.BackgroundColor = null;

                    presentation.Save();
                    Assert.Empty(presentation.ValidateDocument());
                }

                using (PresentationDocument document = PresentationDocument.Open(filePath, false)) {
                    SlidePart slidePart = document.PresentationPart!.SlideParts.First();
                    DocumentFormat.OpenXml.Presentation.BackgroundProperties properties =
                        slidePart.Slide.CommonSlideData!.Background!.BackgroundProperties!;

                    Assert.NotNull(properties.GetFirstChild<A.BlipFill>());
                    Assert.Null(properties.GetFirstChild<A.SolidFill>());
                    Assert.Single(slidePart.ImageParts);
                }
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void ReplacingBackgroundImage_RemovesOrphanedImagePart() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            string originalImage = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Images", "BackgroundImage.png");
            string replacementImage = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Images", "Kulek.jpg");

            try {
                using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                    PowerPointSlide slide = presentation.AddSlide();
                    slide.SetBackgroundImage(originalImage);
                    slide.SetBackgroundImage(replacementImage);

                    presentation.Save();
                    Assert.Empty(presentation.ValidateDocument());
                }

                using (PresentationDocument document = PresentationDocument.Open(filePath, false)) {
                    SlidePart slidePart = document.PresentationPart!.SlideParts.First();
                    ImagePart[] imageParts = slidePart.ImageParts.ToArray();

                    Assert.Single(imageParts);
                    Assert.Equal("image/jpeg", imageParts[0].ContentType);
                }
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void ClearingBackgroundImage_RemovesOrphanedImagePart() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            string imagePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Images", "BackgroundImage.png");

            try {
                using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                    PowerPointSlide slide = presentation.AddSlide();
                    slide.SetBackgroundImage(imagePath);
                    slide.ClearBackgroundImage();

                    presentation.Save();
                    Assert.Empty(presentation.ValidateDocument());
                }

                using (PresentationDocument document = PresentationDocument.Open(filePath, false)) {
                    SlidePart slidePart = document.PresentationPart!.SlideParts.First();
                    DocumentFormat.OpenXml.Presentation.Background? background =
                        slidePart.Slide.CommonSlideData!.Background;

                    Assert.Empty(slidePart.ImageParts);
                    Assert.True(background == null ||
                        background.BackgroundProperties?.GetFirstChild<A.BlipFill>() == null);
                }
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void NotesText_ReadsAllParagraphsAndRuns() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");

            try {
                using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                    PowerPointSlide slide = presentation.AddSlide();
                    slide.Notes.Text = "seed";
                    presentation.Save();
                }

                using (PresentationDocument document = PresentationDocument.Open(filePath, true)) {
                    DocumentFormat.OpenXml.Presentation.Shape notesShape = GetNotesTextShape(document);
                    notesShape.TextBody!.RemoveAllChildren<A.Paragraph>();
                    notesShape.TextBody.Append(
                        new A.Paragraph(
                            new A.Run(new A.Text("Alpha")),
                            new A.Run(new A.Text(" Beta"))),
                        new A.Paragraph(
                            new A.Run(new A.Text("Gamma")),
                            new A.Break(),
                            new A.Run(new A.Text("Delta"))));
                    document.PresentationPart!.SlideParts.First().NotesSlidePart!.NotesSlide!.Save();
                }

                using (PowerPointPresentation presentation = PowerPointPresentation.Open(filePath)) {
                    Assert.Equal($"Alpha Beta{Environment.NewLine}Gamma{Environment.NewLine}Delta",
                        presentation.Slides.Single().Notes.Text);
                }
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void NotesText_SetterReplacesExistingParagraphStructure() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");

            try {
                using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                    PowerPointSlide slide = presentation.AddSlide();
                    slide.Notes.Text = "Old";
                    presentation.Save();
                }

                using (PowerPointPresentation presentation = PowerPointPresentation.Open(filePath)) {
                    presentation.Slides.Single().Notes.Text = "First line\r\nSecond line";
                    presentation.Save();
                }

                using (PresentationDocument document = PresentationDocument.Open(filePath, false)) {
                    DocumentFormat.OpenXml.Presentation.Shape notesShape = GetNotesTextShape(document);
                    A.Paragraph[] paragraphs = notesShape.TextBody!.Elements<A.Paragraph>().ToArray();

                    Assert.Equal(2, paragraphs.Length);
                    Assert.Equal("First line", ReadParagraphText(paragraphs[0]));
                    Assert.Equal("Second line", ReadParagraphText(paragraphs[1]));
                }
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        private static DocumentFormat.OpenXml.Presentation.Shape GetNotesTextShape(PresentationDocument document) {
            return document.PresentationPart!.SlideParts.First().NotesSlidePart!.NotesSlide!.CommonSlideData!.ShapeTree!
                .Elements<DocumentFormat.OpenXml.Presentation.Shape>()
                .First(shape =>
                    shape.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?.GetFirstChild<PlaceholderShape>()?.Type?.Value ==
                    PlaceholderValues.Body);
        }

        private static string ReadParagraphText(A.Paragraph paragraph) {
            StringBuilder builder = new();
            foreach (OpenXmlElement child in paragraph.ChildElements) {
                switch (child) {
                    case A.Run run:
                        builder.Append(run.Text?.Text ?? string.Empty);
                        break;
                    case A.Break:
                        builder.AppendLine();
                        break;
                    case A.Field field:
                        builder.Append(field.Text?.Text ?? string.Empty);
                        break;
                }
            }

            return builder.ToString();
        }
    }
}
