using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.Presentation;
using OfficeIMO.Drawing;
using OfficeIMO.PowerPoint;
using OfficeIMO.Pdf;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.Tests {
    public partial class PowerPointImageExportTests {
        private static readonly IReadOnlyList<PowerPointImageFixtureBaseline> RealWorldFixtureBaselines = new[] {
            new PowerPointImageFixtureBaseline(
                "officeimo-powerpoint-image-dashboard",
                CreateDashboardFixture,
                new[] {
                    "Executive Summary",
                    "Revenue",
                    "$42M",
                    "Retention",
                    "96%",
                    "Risk",
                    "Low",
                    "North",
                    "12.4",
                    "South",
                    "9.8",
                    "East",
                    "11.6",
                    "West",
                    "8.2"
                },
                new[] {
                    OfficeColor.FromRgb(15, 23, 42),
                    OfficeColor.FromRgb(59, 130, 246),
                    OfficeColor.FromRgb(16, 185, 129),
                    OfficeColor.FromRgb(245, 158, 11),
                    OfficeColor.FromRgb(99, 102, 241)
                },
                4000),
            new PowerPointImageFixtureBaseline(
                "officeimo-powerpoint-image-process-diagram",
                CreateProcessDiagramFixture,
                new[] { "Discovery", "Design", "Render", "Validate", "OpenXML", "Drawing", "only" },
                new[] {
                    OfficeColor.FromRgb(30, 64, 175),
                    OfficeColor.FromRgb(34, 197, 94),
                    OfficeColor.FromRgb(234, 88, 12),
                    OfficeColor.FromRgb(168, 85, 247)
                },
                10000)
        };

        [Theory]
        [MemberData(nameof(GetRealWorldFixtureBaselines))]
        public void PowerPointImageExportRealWorldFixturesPreserveTextColorsAndDimensions(PowerPointImageFixtureBaseline baseline) {
            using PowerPointDocument presentation = baseline.CreateDocument();
            PowerPointSlide slide = presentation.Slides[0];

            PowerPointSlideVisualInfo snapshot = slide.CreateVisualSnapshot();
            OfficeImageExportResult png = slide.ExportImage(OfficeImageExportFormat.Png);
            OfficeImageExportResult svg = slide.ExportImage(OfficeImageExportFormat.Svg);

            AssertFixtureDiagnostics(snapshot.Diagnostics);
            AssertFixtureDiagnostics(png.Diagnostics);
            AssertFixtureDiagnostics(svg.Diagnostics);
            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? image), baseline.Name + " PNG output is not decodable by OfficeIMO.Drawing.");
            Assert.Equal((int)Math.Round(snapshot.Width), image!.Width);
            Assert.Equal((int)Math.Round(snapshot.Height), image.Height);
            Assert.True(CountNonBackgroundPixels(image, OfficeColor.White) >= baseline.ExpectedMinimumNonBackgroundPixels, baseline.Name + " rendered as blank or near-blank PNG.");

            string svgText = Encoding.UTF8.GetString(svg.Bytes);
            Assert.Contains("<svg", svgText, StringComparison.Ordinal);
            string snapshotText = GetSnapshotPlainText(snapshot);
            foreach (string expectedText in baseline.ExpectedTexts) {
                Assert.Contains(expectedText, snapshotText, StringComparison.Ordinal);
            }

            foreach (OfficeImageExportFormat format in new[] { OfficeImageExportFormat.Png, OfficeImageExportFormat.Jpeg, OfficeImageExportFormat.Svg }) {
                AssertNativeImageRoute(presentation, format);
            }

            foreach (OfficeColor expectedColor in baseline.ExpectedRasterColors) {
                Assert.True(CountPixelsNear(image, expectedColor) > 20, baseline.Name + " lost expected color " + expectedColor + " in PNG output.");
            }
        }

        private static void AssertFixtureDiagnostics(IEnumerable<OfficeImageExportDiagnostic> diagnostics) {
            OfficeImageExportDiagnostic[] all = diagnostics.ToArray();
            // These compact labels allow visible overflow. Any clipping or other loss still fails below.
            foreach (OfficeImageExportDiagnostic diagnostic in all.Where(item => item.Code == "text-box-overflow" || item.Code == OfficeDiagnosticCodes.PowerpointTextOverflow)) {
                Assert.Equal(OfficeConversionLossKind.None, diagnostic.LossKind);
                Assert.Equal(OfficeImageExportDiagnosticSeverity.Warning, diagnostic.Severity);
            }
            AssertNoUnexpectedDiagnostics(all.Where(item => item.Code != "text-box-overflow" && item.Code != OfficeDiagnosticCodes.PowerpointTextOverflow));
        }

        public static IEnumerable<object[]> GetRealWorldFixtureBaselines() =>
            RealWorldFixtureBaselines.Select(item => new object[] { item });

        private static PowerPointDocument CreateDashboardFixture() {
            var stream = new MemoryStream();
            PowerPointDocument presentation = PowerPointDocument.Create(stream);
            presentation.SlideSize.SetSize(OfficeIMO.OfficeLength.Points(320),OfficeIMO.OfficeLength.Points(180));
            PowerPointSlide slide = presentation.AddSlide();
            slide.BackgroundColor = "FFFFFF";

            PowerPointTextBox title = slide.AddTextBox("Executive Summary",global::OfficeIMO.OfficeLength.Points(18),global::OfficeIMO.OfficeLength.Points(12),global::OfficeIMO.OfficeLength.Points(260),global::OfficeIMO.OfficeLength.Points(24));
            title.FontSize = global::OfficeIMO.OfficeLength.Points(18);
            title.Color = "0F172A";

            AddMetricCard(slide, "Revenue", "$42M", "3B82F6", 18, 42);
            AddMetricCard(slide, "Retention", "96%", "10B981", 116, 42);
            AddMetricCard(slide, "Risk", "Low", "F59E0B", 214, 42);

            PowerPointTable table = slide.AddTable(4,2,global::OfficeIMO.OfficeLength.Points(18),global::OfficeIMO.OfficeLength.Points(104),global::OfficeIMO.OfficeLength.Points(132),global::OfficeIMO.OfficeLength.Points(54));
            table.SetColumnWidths(global::OfficeIMO.OfficeLength.Points(66),global::OfficeIMO.OfficeLength.Points(66));
            table.SetRowHeights(global::OfficeIMO.OfficeLength.Points(13.5),global::OfficeIMO.OfficeLength.Points(13.5),global::OfficeIMO.OfficeLength.Points(13.5),global::OfficeIMO.OfficeLength.Points(13.5));
            string[] regions = { "North", "South", "East", "West" };
            string[] values = { "12.4", "9.8", "11.6", "8.2" };
            for (int row = 0; row < regions.Length; row++) {
                PowerPointTableCell region = table.GetCell(row, 0);
                region.Text = regions[row];
                region.FontSize = global::OfficeIMO.OfficeLength.Points(8);
                region.FillColor = row % 2 == 0 ? "F8FAFC" : "EEF2FF";

                PowerPointTableCell value = table.GetCell(row, 1);
                value.Text = values[row];
                value.FontSize = global::OfficeIMO.OfficeLength.Points(8);
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
            slide.AddImage(image, OfficeImageFormat.Bmp, OfficeIMO.OfficeLength.Points(276), OfficeIMO.OfficeLength.Points(16), OfficeIMO.OfficeLength.Points(24), OfficeIMO.OfficeLength.Points(24));

            return presentation;
        }

        private static PowerPointDocument CreateProcessDiagramFixture() {
            var stream = new MemoryStream();
            PowerPointDocument presentation = PowerPointDocument.Create(stream);
            presentation.SlideSize.SetSize(OfficeIMO.OfficeLength.Points(320),OfficeIMO.OfficeLength.Points(180));
            PowerPointSlide slide = presentation.AddSlide();
            slide.SetBackgroundGradient("EFF6FF", "F8FAFC", 0D);

            AddProcessNode(slide, "Discovery", "1E40AF", 20, 58);
            AddProcessNode(slide, "Design", "22C55E", 96, 58);
            AddProcessNode(slide, "Render", "EA580C", 172, 58);
            AddProcessNode(slide, "Validate", "A855F7", 248, 58);

            AddConnector(slide, 86, 76, 10, 0);
            AddConnector(slide, 162, 76, 10, 0);
            AddConnector(slide, 238, 76, 10, 0);

            PowerPointTextBox note = slide.AddTextBox("OpenXML + Drawing only",global::OfficeIMO.OfficeLength.Points(56),global::OfficeIMO.OfficeLength.Points(120),global::OfficeIMO.OfficeLength.Points(208),global::OfficeIMO.OfficeLength.Points(20));
            note.FontSize = global::OfficeIMO.OfficeLength.Points(12);
            note.Color = "0F172A";

            return presentation;
        }

        private static void AddMetricCard(PowerPointSlide slide, string label, string value, string color, double left, double top) {
            PowerPointAutoShape card = slide.AddRectangle(OfficeIMO.OfficeLength.Points(left), OfficeIMO.OfficeLength.Points(top), OfficeIMO.OfficeLength.Points(86), OfficeIMO.OfficeLength.Points(46));
            card.FillColor = "F8FAFC";
            card.OutlineColor = color;
            card.OutlineWidth = global::OfficeIMO.OfficeLength.Points(1.5D);

            PowerPointAutoShape accent = slide.AddRectangle(OfficeIMO.OfficeLength.Points(left), OfficeIMO.OfficeLength.Points(top), OfficeIMO.OfficeLength.Points(4), OfficeIMO.OfficeLength.Points(46));
            accent.FillColor = color;
            accent.OutlineColor = color;
            accent.OutlineWidth = global::OfficeIMO.OfficeLength.Points(0D);

            PowerPointTextBox labelBox = slide.AddTextBox(label,global::OfficeIMO.OfficeLength.Points(left + 8),global::OfficeIMO.OfficeLength.Points(top + 7),global::OfficeIMO.OfficeLength.Points(68),global::OfficeIMO.OfficeLength.Points(12));
            labelBox.FontSize = global::OfficeIMO.OfficeLength.Points(8);
            labelBox.Color = "334155";

            PowerPointTextBox valueBox = slide.AddTextBox(value,global::OfficeIMO.OfficeLength.Points(left + 8),global::OfficeIMO.OfficeLength.Points(top + 22),global::OfficeIMO.OfficeLength.Points(68),global::OfficeIMO.OfficeLength.Points(16));
            valueBox.FontSize = global::OfficeIMO.OfficeLength.Points(14);
            valueBox.Color = color;
        }

        private static void AddProcessNode(PowerPointSlide slide, string label, string color, double left, double top) {
            PowerPointAutoShape node = slide.AddShape(OfficePresetShapeType.RoundRectangle, OfficeIMO.OfficeLength.Points(left), OfficeIMO.OfficeLength.Points(top), OfficeIMO.OfficeLength.Points(66), OfficeIMO.OfficeLength.Points(36));
            node.FillColor = color;
            node.OutlineColor = "0F172A";
            node.OutlineWidth = global::OfficeIMO.OfficeLength.Points(1D);

            PowerPointTextBox text = slide.AddTextBox(label,global::OfficeIMO.OfficeLength.Points(left + 4),global::OfficeIMO.OfficeLength.Points(top + 11),global::OfficeIMO.OfficeLength.Points(58),global::OfficeIMO.OfficeLength.Points(12));
            text.TextMarginLeft = text.TextMarginRight = global::OfficeIMO.OfficeLength.Points(0D);
            text.TextMarginTop = text.TextMarginBottom = global::OfficeIMO.OfficeLength.Points(0D);
            text.FontSize = global::OfficeIMO.OfficeLength.Points(8);
            text.Color = "FFFFFF";
        }

        private static void AddConnector(PowerPointSlide slide, double left, double top, double width, double height) {
            PowerPointAutoShape connector = slide.AddShape(OfficePresetShapeType.Line, OfficeIMO.OfficeLength.Points(left), OfficeIMO.OfficeLength.Points(top), OfficeIMO.OfficeLength.Points(width), OfficeIMO.OfficeLength.Points(height));
            connector.OutlineColor = "0F172A";
            connector.OutlineWidth = global::OfficeIMO.OfficeLength.Points(1.2D);
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

        private static string GetSnapshotPlainText(PowerPointSlideVisualInfo snapshot) =>
            string.Concat(GetSnapshotTextSegments(snapshot).Select(segment => segment.Text));

        private static IEnumerable<(PowerPointTextSpanInfo Span, string Text)> GetSnapshotTextSegments(PowerPointSlideVisualInfo snapshot) {
            PowerPointTextSpanInfo? previous = null;
            foreach (PowerPointTextSpanInfo span in snapshot.TextSpans) {
                string text = System.Text.RegularExpressions.Regex.Replace(span.Text, @"\s+", " ");
                if (text.Length == 0) continue;
                // Preserve run boundaries within a line; a wrapped line separates logical words.
                if (previous != null && Math.Abs(previous.Y - span.Y) > 0.01D
                    && previous.Text.Length > 0 && !char.IsWhiteSpace(previous.Text[previous.Text.Length - 1])
                    && !char.IsWhiteSpace(text[0])) text = " " + text;
                yield return (span, text);
                previous = span;
            }
        }

        public sealed class PowerPointImageFixtureBaseline {
            public PowerPointImageFixtureBaseline(
                string name,
                Func<PowerPointDocument> createPresentation,
                IReadOnlyList<string> expectedTexts,
                IReadOnlyList<OfficeColor> expectedRasterColors,
                int expectedMinimumNonBackgroundPixels) {
                Name = name;
                CreateDocument = createPresentation;
                ExpectedTexts = expectedTexts;
                ExpectedRasterColors = expectedRasterColors;
                ExpectedMinimumNonBackgroundPixels = expectedMinimumNonBackgroundPixels;
            }

            public string Name { get; }

            public Func<PowerPointDocument> CreateDocument { get; }

            public IReadOnlyList<string> ExpectedTexts { get; }

            public IReadOnlyList<OfficeColor> ExpectedRasterColors { get; }

            public int ExpectedMinimumNonBackgroundPixels { get; }
        }
    }
}
