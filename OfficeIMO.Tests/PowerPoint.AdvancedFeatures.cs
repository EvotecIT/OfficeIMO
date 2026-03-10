using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Packaging;
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
    }
}
