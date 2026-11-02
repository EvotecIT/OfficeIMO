using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Drawing.Charts;
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
                using PowerPointDocument presentation = PowerPointDocument.Create(filePath);
                PowerPointSlide slide = presentation.AddSlide();
                PowerPointAutoShape original = slide.AddRectangle(OfficeIMO.OfficeLength.Emus(1000), OfficeIMO.OfficeLength.Emus(2000), OfficeIMO.OfficeLength.Emus(3000), OfficeIMO.OfficeLength.Emus(4000));

                PowerPointShape duplicate = slide.DuplicateShape(original, OfficeIMO.OfficeLength.Emus(500), OfficeIMO.OfficeLength.Emus(600));

                Assert.Equal(2, slide.Shapes.Count);
                Assert.Equal(original.Width.ToEmus(), duplicate.Width.ToEmus());
                Assert.Equal(original.Height.ToEmus(), duplicate.Height.ToEmus());
                Assert.Equal(original.Left.ToEmus() + 500, duplicate.Left.ToEmus());
                Assert.Equal(original.Top.ToEmus() + 600, duplicate.Top.ToEmus());
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
                using PowerPointDocument presentation = PowerPointDocument.Create(filePath);
                PowerPointSlide slide = presentation.AddSlide();
                PowerPointAutoShape a = slide.AddRectangle(OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(1000), OfficeIMO.OfficeLength.Emus(1000));
                PowerPointAutoShape b = slide.AddRectangle(OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(1000), OfficeIMO.OfficeLength.Emus(1000));
                PowerPointAutoShape c = slide.AddRectangle(OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(1000), OfficeIMO.OfficeLength.Emus(1000));

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
                using PowerPointDocument presentation = PowerPointDocument.Create(filePath);
                PowerPointSlide slide = presentation.AddSlide();
                slide.AddRectangle(OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(1000), OfficeIMO.OfficeLength.Emus(1000));
                slide.AddRectangle(OfficeIMO.OfficeLength.Emus(1500), OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(1000), OfficeIMO.OfficeLength.Emus(1000));

                PowerPointGroupShape group = slide.GroupShapes(slide.Shapes);
                PowerPointGroupShape duplicate = Assert.IsType<PowerPointGroupShape>(slide.DuplicateShape(group, OfficeIMO.OfficeLength.Emus(500), OfficeIMO.OfficeLength.Emus(500)));

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

        [Fact]
        public void DuplicateChart_CreatesIndependentChartPart() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");

            try {
                using (PowerPointDocument presentation = PowerPointDocument.Create(filePath)) {
                    PowerPointSlide slide = presentation.AddSlide();
                    PowerPointChart original = slide.AddChart();
                    PowerPointChart duplicate = Assert.IsType<PowerPointChart>(slide.DuplicateShape(original, OfficeIMO.OfficeLength.Emus(250000), OfficeIMO.OfficeLength.Emus(0)));

                    duplicate.UpdateData(new PowerPointChartData(
                        new[] { "Jan", "Feb" },
                        new[] { new PowerPointChartSeries("Duplicate", new[] { 10d, 20d }) }));

                    presentation.Save();
                }

                using (PresentationDocument document = PresentationDocument.Open(filePath, false)) {
                    SlidePart slidePart = document.PresentationPart!.SlideParts.First();
                    GraphicFrame[] chartFrames = slidePart.Slide.CommonSlideData!.ShapeTree!
                        .Elements<GraphicFrame>()
                        .Where(frame => frame.Graphic?.GraphicData?.GetFirstChild<ChartReference>() != null)
                        .ToArray();

                    Assert.Equal(2, chartFrames.Length);

                    string[] relationshipIds = chartFrames
                        .Select(frame => frame.Graphic!.GraphicData!.GetFirstChild<ChartReference>()!.Id!.Value!)
                        .ToArray();
                    Assert.Equal(2, relationshipIds.Distinct(StringComparer.Ordinal).Count());

                    ChartPart originalChartPart = (ChartPart)slidePart.GetPartById(relationshipIds[0]);
                    ChartPart duplicateChartPart = (ChartPart)slidePart.GetPartById(relationshipIds[1]);

                    List<string> originalSeries = GetSeriesNames(originalChartPart);
                    List<string> duplicateSeries = GetSeriesNames(duplicateChartPart);

                    Assert.Equal(new[] { "Series 1", "Series 2", "Series 3" }, originalSeries);
                    Assert.Equal(new[] { "Duplicate" }, duplicateSeries);
                }
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        private static List<string> GetSeriesNames(ChartPart chartPart) {
            return chartPart.ChartSpace!.Descendants<SeriesText>()
                .Select(series => series.StringReference?.StringCache?.Elements<StringPoint>()
                    .Select(point => point.NumericValue?.Text)
                    .LastOrDefault(text => !string.IsNullOrWhiteSpace(text))
                    ?? series.InnerText)
                .Where(text => !string.IsNullOrWhiteSpace(text))
                .ToList();
        }
    }
}
