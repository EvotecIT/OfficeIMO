using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.Presentation;
using OfficeIMO.Drawing;
using OfficeIMO.PowerPoint;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.Tests {
    public partial class PowerPointImageExportTests {
        private static readonly IReadOnlyList<PowerPointImageFixtureBaseline> RealWorldFixtureBaselines = new[] {
            new PowerPointImageFixtureBaseline(
                "officeimo-powerpoint-image-dashboard",
                CreateDashboardFixture,
                new[] { "Executive Summary", "Revenue", "Retention", "North", "South", "East", "West" },
                new[] {
                    OfficeColor.FromRgb(15, 23, 42),
                    OfficeColor.FromRgb(59, 130, 246),
                    OfficeColor.FromRgb(16, 185, 129),
                    OfficeColor.FromRgb(245, 158, 11),
                    OfficeColor.FromRgb(99, 102, 241)
                },
                18,
                4000),
            new PowerPointImageFixtureBaseline(
                "officeimo-powerpoint-image-process-diagram",
                CreateProcessDiagramFixture,
                new[] { "Discovery", "Design", "Render", "Validate" },
                new[] {
                    OfficeColor.FromRgb(30, 64, 175),
                    OfficeColor.FromRgb(34, 197, 94),
                    OfficeColor.FromRgb(234, 88, 12),
                    OfficeColor.FromRgb(168, 85, 247)
                },
                13,
                10000)
        };

        [Fact]
        public void PowerPointImageExportRealWorldFixtureManifestCoversApprovedFixtureBaselines() {
            string[] names = RealWorldFixtureBaselines
                .Select(item => item.Name)
                .OrderBy(name => name, StringComparer.Ordinal)
                .ToArray();

            Assert.Equal(
                new[] {
                    "officeimo-powerpoint-image-dashboard",
                    "officeimo-powerpoint-image-process-diagram"
                },
                names);
        }

        [Theory]
        [MemberData(nameof(GetRealWorldFixtureBaselines))]
        public void PowerPointImageExportRealWorldFixturesRenderThroughSharedDrawing(PowerPointImageFixtureBaseline baseline) {
            using PowerPointPresentation presentation = baseline.CreatePresentation();
            PowerPointSlide slide = presentation.Slides[0];

            PowerPointSlideVisualSnapshot snapshot = slide.CreateVisualSnapshot();
            OfficeImageExportResult png = slide.ExportImage(OfficeImageExportFormat.Png);
            OfficeImageExportResult svg = slide.ExportImage(OfficeImageExportFormat.Svg);

            Assert.Empty(snapshot.Diagnostics);
            Assert.Empty(png.Diagnostics);
            Assert.Empty(svg.Diagnostics);
            Assert.True(snapshot.Drawing.Elements.Count >= baseline.ExpectedMinimumDrawingElements, baseline.Name + " lost Drawing snapshot coverage.");
            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? image), baseline.Name + " PNG output is not decodable by OfficeIMO.Drawing.");
            Assert.Equal((int)Math.Round(snapshot.Width), image!.Width);
            Assert.Equal((int)Math.Round(snapshot.Height), image.Height);
            Assert.True(CountNonBackgroundPixels(image, OfficeColor.White) >= baseline.ExpectedMinimumNonBackgroundPixels, baseline.Name + " rendered as blank or near-blank PNG.");

            string svgText = Encoding.UTF8.GetString(svg.Bytes);
            Assert.Contains("<svg", svgText, StringComparison.Ordinal);
            string snapshotText = GetSnapshotPlainText(snapshot);
            foreach (string expectedText in baseline.ExpectedSvgTexts) {
                Assert.Contains(expectedText, snapshotText, StringComparison.Ordinal);
            }

            foreach (OfficeColor expectedColor in baseline.ExpectedRasterColors) {
                Assert.True(CountPixelsNear(image, expectedColor) > 20, baseline.Name + " lost expected color " + expectedColor + " in PNG output.");
            }
        }

        public static IEnumerable<object[]> GetRealWorldFixtureBaselines() =>
            RealWorldFixtureBaselines.Select(item => new object[] { item });

        private static PowerPointPresentation CreateDashboardFixture() {
            var stream = new MemoryStream();
            PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
            presentation.SlideSize.SetSizePoints(320, 180);
            PowerPointSlide slide = presentation.Slides[0];
            slide.BackgroundColor = "FFFFFF";

            PowerPointTextBox title = slide.AddTextBoxPoints("Executive Summary", 18, 12, 190, 22);
            title.FontSize = 18;
            title.Color = "0F172A";

            AddMetricCard(slide, "Revenue", "$42M", "3B82F6", 18, 42);
            AddMetricCard(slide, "Retention", "96%", "10B981", 116, 42);
            AddMetricCard(slide, "Risk", "Low", "F59E0B", 214, 42);

            PowerPointTable table = slide.AddTablePoints(4, 2, 18, 104, 132, 54);
            table.SetColumnWidthsPoints(66, 66);
            table.SetRowHeightsPoints(13.5, 13.5, 13.5, 13.5);
            string[] regions = { "North", "South", "East", "West" };
            string[] values = { "12.4", "9.8", "11.6", "8.2" };
            for (int row = 0; row < regions.Length; row++) {
                PowerPointTableCell region = table.GetCell(row, 0);
                region.Text = regions[row];
                region.FontSize = 8;
                region.FillColor = row % 2 == 0 ? "F8FAFC" : "EEF2FF";

                PowerPointTableCell value = table.GetCell(row, 1);
                value.Text = values[row];
                value.FontSize = 8;
                value.FillColor = row % 2 == 0 ? "F8FAFC" : "EEF2FF";
            }

            var data = new PowerPointChartData(
                new[] { "Q1", "Q2", "Q3", "Q4" },
                new[] {
                    new PowerPointChartSeries("North", new[] { 8D, 9D, 10D, 12D }),
                    new PowerPointChartSeries("South", new[] { 7D, 8D, 9D, 10D })
                });
            slide.AddChartPoints(data, 166, 96, 132, 62);

            using var image = new MemoryStream(CreateBmp24(
                2,
                2,
                new[] {
                    OfficeColor.FromRgb(99, 102, 241),
                    OfficeColor.White,
                    OfficeColor.White,
                    OfficeColor.FromRgb(99, 102, 241)
                },
                topDown: true));
            slide.AddPicturePoints(image, ImagePartType.Bmp, 276, 16, 24, 24);

            return presentation;
        }

        private static PowerPointPresentation CreateProcessDiagramFixture() {
            var stream = new MemoryStream();
            PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
            presentation.SlideSize.SetSizePoints(320, 180);
            PowerPointSlide slide = presentation.Slides[0];
            slide.SetBackgroundGradient("EFF6FF", "F8FAFC", 0D);

            AddProcessNode(slide, "Discovery", "1E40AF", 20, 58);
            AddProcessNode(slide, "Design", "22C55E", 96, 58);
            AddProcessNode(slide, "Render", "EA580C", 172, 58);
            AddProcessNode(slide, "Validate", "A855F7", 248, 58);

            AddConnector(slide, 78, 76, 18, 0);
            AddConnector(slide, 154, 76, 18, 0);
            AddConnector(slide, 230, 76, 18, 0);

            PowerPointTextBox note = slide.AddTextBoxPoints("OpenXML + Drawing only", 72, 120, 176, 20);
            note.FontSize = 12;
            note.Color = "0F172A";

            return presentation;
        }

        private static void AddMetricCard(PowerPointSlide slide, string label, string value, string color, double left, double top) {
            PowerPointAutoShape card = slide.AddRectanglePoints(left, top, 86, 46);
            card.FillColor = "F8FAFC";
            card.OutlineColor = color;
            card.OutlineWidthPoints = 1.5D;

            PowerPointAutoShape accent = slide.AddRectanglePoints(left, top, 4, 46);
            accent.FillColor = color;
            accent.OutlineColor = color;
            accent.OutlineWidthPoints = 0D;

            PowerPointTextBox labelBox = slide.AddTextBoxPoints(label, left + 8, top + 7, 68, 12);
            labelBox.FontSize = 8;
            labelBox.Color = "334155";

            PowerPointTextBox valueBox = slide.AddTextBoxPoints(value, left + 8, top + 22, 68, 16);
            valueBox.FontSize = 14;
            valueBox.Color = color;
        }

        private static void AddProcessNode(PowerPointSlide slide, string label, string color, double left, double top) {
            PowerPointAutoShape node = slide.AddShapePoints(A.ShapeTypeValues.RoundRectangle, left, top, 54, 36);
            node.FillColor = color;
            node.OutlineColor = "0F172A";
            node.OutlineWidthPoints = 1D;

            PowerPointTextBox text = slide.AddTextBoxPoints(label, left + 5, top + 11, 44, 12);
            text.FontSize = 8;
            text.Color = "FFFFFF";
        }

        private static void AddConnector(PowerPointSlide slide, double left, double top, double width, double height) {
            PowerPointAutoShape connector = slide.AddShapePoints(A.ShapeTypeValues.Line, left, top, width, height);
            connector.OutlineColor = "0F172A";
            connector.OutlineWidthPoints = 1.2D;
        }

        private static int CountNonBackgroundPixels(OfficeRasterImage image, OfficeColor background) {
            int count = 0;
            for (int y = 0; y < image.Height; y++) {
                for (int x = 0; x < image.Width; x++) {
                    OfficeColor pixel = image.GetPixel(x, y);
                    if (Math.Abs(pixel.R - background.R) > 8 ||
                        Math.Abs(pixel.G - background.G) > 8 ||
                        Math.Abs(pixel.B - background.B) > 8 ||
                        Math.Abs(pixel.A - background.A) > 8) {
                        count++;
                    }
                }
            }

            return count;
        }

        private static string GetSnapshotPlainText(PowerPointSlideVisualSnapshot snapshot) {
            var builder = new StringBuilder();
            foreach (OfficeDrawingElement element in snapshot.Drawing.Elements) {
                switch (element) {
                    case OfficeDrawingText text:
                        builder.AppendLine(text.Text);
                        break;
                    case OfficeDrawingRichText richText:
                        builder.AppendLine(richText.PlainText);
                        break;
                }
            }

            return builder.ToString();
        }

        public sealed record PowerPointImageFixtureBaseline(
            string Name,
            Func<PowerPointPresentation> CreatePresentation,
            IReadOnlyList<string> ExpectedSvgTexts,
            IReadOnlyList<OfficeColor> ExpectedRasterColors,
            int ExpectedMinimumDrawingElements,
            int ExpectedMinimumNonBackgroundPixels);
    }
}
