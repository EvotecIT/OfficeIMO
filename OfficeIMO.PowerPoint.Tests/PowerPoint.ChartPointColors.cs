using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Drawing;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.Pdf;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using P = DocumentFormat.OpenXml.Presentation;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Tests {
    public class PowerPointChartPointColorsTests {
        [Fact]
        public void NotesChartUsesNotesColorMapOverride() {
            using PowerPointDocument presentation = PowerPointDocument.Create();
            PowerPointSlide slide = presentation.AddSlide();
            slide.Notes.Text = "Notes";
            NotesSlidePart notes = slide.SlidePart.NotesSlidePart!;
            ChartPart chart = notes.AddNewPart<ChartPart>();
            chart.ChartSpace = new C.ChartSpace();
            notes.NotesSlide!.AddChild(new P.ColorMapOverride(new A.OverrideColorMapping {
                Accent1 = A.ColorSchemeIndexValues.Accent2
            }), true);

            var map = OfficeIMO.OpenXml.Internal.OfficeOpenXmlThemeColorResolver.ResolveChartColorMap(chart);
            Assert.Equal("accent2", OfficeIMO.OpenXml.Internal.OfficeOpenXmlThemeColorResolver.MapSchemeColor("accent1", map));
        }

        private static readonly OfficeColor?[] Colors = {
            OfficeColor.FromRgb(18, 52, 86), null, OfficeColor.FromRgb(254, 220, 186)
        };

        public static IEnumerable<object[]> CategoryKinds() => PowerPointSharedChartCases.Kinds.Where(kind => kind != OfficeChartType.Bubble)
            .Select(kind => new object[] { kind });

        [Theory]
        [MemberData(nameof(CategoryKinds))]
        public void PointColors_SurviveNativeSaveReopenAndDataUpdate(OfficeChartType kind) {
            using PowerPointDocument presentation = PowerPointDocument.Create();
            PowerPointChart chart = presentation.AddSlide().AddChartCm(kind, CreateData(kind, Colors), 1, 1, 20, 10);
            AssertColors(chart, Colors);
            chart.UpdateData(CreateData(kind, new OfficeColor?[] { Colors[2], null, Colors[0] }));
            AssertColors(chart, new OfficeColor?[] { Colors[2], null, Colors[0] });
            using var bytes = new MemoryStream(presentation.ToBytes());
            using PowerPointDocument reopened = PowerPointDocument.Load(bytes);
            AssertColors(Assert.Single(reopened.Slides.Single().Charts), new OfficeColor?[] { Colors[2], null, Colors[0] });
            Assert.Empty(reopened.ValidateDocument());
        }

        [Fact]
        public void PointColors_ExplicitNullClearsAnOldNativePointFill() {
            using PowerPointDocument presentation = PowerPointDocument.Create();
            PowerPointChart chart = presentation.AddSlide().AddChart(OfficeChartType.Pie,
                CreateData(OfficeChartType.Pie, Colors));
            chart.UpdateData(CreateData(OfficeChartType.Pie,
                new OfficeColor?[] { null, Colors[1], Colors[2] }));
            C.PieChartSeries series = presentation.Slides.Single().SlidePart.ChartParts.Single()
                .ChartSpace!.Descendants<C.PieChartSeries>().Single();
            C.DataPoint point = series.Elements<C.DataPoint>().Single(item => item.Index?.Val?.Value == 0);
            Assert.Null(point.ChartShapeProperties?.GetFirstChild<A.SolidFill>());
            Assert.Null(point.ChartShapeProperties?.GetFirstChild<A.PatternFill>());
            AssertColors(chart, new OfficeColor?[] { null, Colors[1], Colors[2] });
        }

        [Fact]
        public void PointColors_ReplaceImportedMarkerFillAndRetainItsOutline() {
            using PowerPointDocument presentation = PowerPointDocument.Create();
            PowerPointChart chart = presentation.AddSlide().AddChart(OfficeChartType.Line,
                CreateData(OfficeChartType.Line, null));
            C.LineChartSeries series = presentation.Slides.Single().SlidePart.ChartParts.Single()
                .ChartSpace!.Descendants<C.LineChartSeries>().Single();
            var outline = new A.Outline(new A.SolidFill(new A.RgbColorModelHex { Val = "123456" }));
            series.AddChild(new C.DataPoint(new C.Index { Val = 0U },
                new C.Marker(new C.ChartShapeProperties(
                    new A.SolidFill(new A.RgbColorModelHex { Val = "FF0000" }), outline))), true);

            OfficeColor replacement = OfficeColor.FromRgb(31, 78, 121);
            chart.UpdateData(CreateData(OfficeChartType.Line,
                new OfficeColor?[] { replacement, null, null }));
            series = presentation.Slides.Single().SlidePart.ChartParts.Single()
                .ChartSpace!.Descendants<C.LineChartSeries>().Single();
            C.DataPoint point = series.Elements<C.DataPoint>().Single(item => item.Index?.Val?.Value == 0);
            Assert.Equal("1F4E79", point.Marker!.ChartShapeProperties!.GetFirstChild<A.SolidFill>()!
                .GetFirstChild<A.RgbColorModelHex>()!.Val!.Value);
            Assert.Equal("123456", point.Marker.ChartShapeProperties.GetFirstChild<A.Outline>()!
                .GetFirstChild<A.SolidFill>()!.GetFirstChild<A.RgbColorModelHex>()!.Val!.Value);
            Assert.Null(point.ChartShapeProperties);
            Assert.Empty(presentation.ValidateDocument());
        }

        [Fact]
        public void PointColors_KeepComboSeriesAlignmentAndPdfStyleOverrides() {
            using PowerPointDocument presentation = PowerPointDocument.Create();
            var data = new OfficeChartData(new[] { "First", "Second", "Third" }, new[] {
                new OfficeSeries("Columns", new double[] { 3, 2, 1 }, null, null, Colors,
                    showMarkers: false, renderKind: OfficeChartType.ColumnClustered),
                new OfficeSeries("Line", new double[] { 1, 2, 3 }, null, null,
                    new OfficeColor?[] { Colors[2], null, Colors[0] }, showMarkers: true,
                    renderKind: OfficeChartType.Line, axisGroup: OfficeChartAxisGroup.Secondary)
            });
            PowerPointChart chart = presentation.AddSlide().AddChartPoints(OfficeChartType.ColumnClustered, data, 20, 20, 600, 320);
            chart.SetChartAreaStyle(fillColor: "ABCDEF");
            Assert.True(chart.TryGetSnapshot(out PowerPointChartSnapshot native));
            OfficeChartSnapshot projected = PowerPointPdfConverterExtensions.CreateOfficeChartSnapshot(native, 600, 320, new PowerPointPdfExportOptions());
            Assert.Equal(Colors, projected.Data.Series[0].PointColors);
            Assert.Equal(data.Series[1].PointColors, projected.Data.Series[1].PointColors);
            Assert.Equal(OfficeChartAxisGroup.Secondary, projected.Data.Series[1].AxisGroup);
            Assert.Same(native.Style, projected.Style);
            var style = new OfficeChartStyle(backgroundColor: OfficeColor.FromRgb(250, 240, 230));
            OfficeChartSnapshot overridden = PowerPointPdfConverterExtensions.CreateOfficeChartSnapshot(native, 400, 200,
                new PowerPointPdfExportOptions { ChartStyle = style });
            Assert.Same(style, overridden.Style);
            Assert.Equal(Colors, overridden.Data.Series[0].PointColors);
            Assert.All(overridden.Data.Series, series => Assert.Null(series.Color));
            using var bytes = new MemoryStream(presentation.ToBytes());
            using PowerPointDocument reopened = PowerPointDocument.Load(bytes);
            Assert.True(reopened.Slides.Single().Charts.Single().TryGetSnapshot(out PowerPointChartSnapshot persisted));
            OfficeChartSnapshot persistedOverride = PowerPointPdfConverterExtensions.CreateOfficeChartSnapshot(
                persisted, 400, 200, new PowerPointPdfExportOptions { ChartStyle = style });
            Assert.All(persistedOverride.Data.Series, series => Assert.Null(series.Color));
        }

        [Fact]
        public void PdfStylePreservesExplicitNativeSeriesColorThatMatchesDefaultPalette() {
            OfficeColor explicitColor = OfficeChartStyle.Default.GetSeriesColor(0);
            using PowerPointDocument presentation = PowerPointDocument.Create();
            PowerPointChart chart = presentation.AddSlide().AddChartPoints(OfficeChartType.ColumnClustered,
                new OfficeChartData(new[] { "A" }, new[] {
                    new OfficeSeries("Explicit", new[] { 3d }, null, explicitColor)
                }), 20, 20, 600, 320);
            Assert.True(chart.TryGetSnapshot(out PowerPointChartSnapshot native));
            var customStyle = new OfficeChartStyle(palette: new[] { OfficeColor.Parse("#123456") });
            OfficeChartSnapshot projected = PowerPointPdfConverterExtensions.CreateOfficeChartSnapshot(native,
                600, 320, new PowerPointPdfExportOptions { ChartStyle = customStyle });
            Assert.Equal(explicitColor, projected.Data.Series[0].Color);
            using var bytes = new MemoryStream(presentation.ToBytes());
            using PowerPointDocument reopened = PowerPointDocument.Load(bytes);
            Assert.True(reopened.Slides.Single().Charts.Single().TryGetSnapshot(out PowerPointChartSnapshot persisted));
            OfficeChartSnapshot persistedOverride = PowerPointPdfConverterExtensions.CreateOfficeChartSnapshot(
                persisted, 600, 320, new PowerPointPdfExportOptions { ChartStyle = customStyle });
            Assert.Equal(explicitColor, persistedOverride.Data.Series[0].Color);
        }

        [Fact]
        public void PointColors_RejectExcessDuplicateAndOutOfRangeOverrides() {
            using PowerPointDocument presentation = PowerPointDocument.Create();
            presentation.AddSlide().AddChartCm(OfficeChartType.Pie, CreateData(OfficeChartType.Pie, null), 1, 1, 20, 10);
            using var bytes = new MemoryStream();
            byte[] package = presentation.ToBytes();
            bytes.Write(package, 0, package.Length);
            bytes.Position = 0;
            using (PresentationDocument native = PresentationDocument.Open(bytes, true)) {
                var part = native.PresentationPart!.SlideParts.Single().ChartParts.Single();
                C.PieChartSeries series = part.ChartSpace.Descendants<C.PieChartSeries>().Single();
                for (int i = 0; i <= PowerPointUtils.MaximumSharedChartPoints; i++)
                    series.Append(new C.DataPoint(new C.Index { Val = i % 2 == 0 ? 0U : uint.MaxValue }));
                part.ChartSpace.Save();
            }
            bytes.Position = 0;
            using PowerPointDocument reopened = PowerPointDocument.Load(bytes);
            Assert.False(Assert.Single(reopened.Slides.Single().Charts).TryGetOfficeSnapshot(out _));
        }

        [Fact]
        public void PointColors_DoughnutLegendUsesTheSameSparseFallbackAsSlices() {
            using PowerPointDocument presentation = PowerPointDocument.Create();
            PowerPointChart chart = presentation.AddSlide().AddChartCm(OfficeChartType.Doughnut,
                CreateData(OfficeChartType.Doughnut, Colors), 1, 1, 20, 10);
            Assert.True(chart.TryGetOfficeSnapshot(out OfficeChartSnapshot snapshot));
            OfficeDrawing drawing = OfficeChartDrawingRenderer.Render(snapshot);
            OfficeColor[] swatches = drawing.Shapes.Where(shape =>
                (shape.Shape.Kind == OfficeShapeKind.Rectangle || shape.Shape.Kind == OfficeShapeKind.Polygon) &&
                shape.Shape.FillColor.HasValue && shape.Shape.Width < 20 && shape.Shape.Height < 20)
                .Select(shape => shape.Shape.FillColor!.Value).ToArray();
            Assert.Equal(new[] { Colors[0]!.Value, OfficeColor.FromRgb(90, 100, 110), Colors[2]!.Value }, swatches);
        }

        [Theory]
        [InlineData(OfficeChartType.ColumnClustered)]
        [InlineData(OfficeChartType.BarClustered)]
        public void PointColors_SparseBarOverridesKeepTheSeriesFill(OfficeChartType kind) {
            using PowerPointDocument presentation = PowerPointDocument.Create();
            PowerPointChart chart = presentation.AddSlide().AddChartCm(kind,
                CreateData(kind, Colors), 1, 1, 20, 10);
            Assert.True(chart.TryGetOfficeSnapshot(out OfficeChartSnapshot snapshot));
            OfficeDrawing drawing = OfficeChartDrawingRenderer.Render(snapshot);
            OfficeColor inherited = OfficeColor.FromRgb(90, 100, 110);
            OfficeDrawingShape[] bars = drawing.Shapes.Where(shape =>
                shape.Shape.Kind == OfficeShapeKind.Rectangle &&
                (shape.Shape.Width > 20D || shape.Shape.Height > 20D) &&
                (shape.Shape.FillColor == inherited || shape.Shape.FillColor == Colors[0] ||
                 shape.Shape.FillColor == Colors[2])).ToArray();
            Assert.Equal(3, bars.Length);
            Assert.Equal(inherited, bars[1].Shape.FillColor);
        }

        [Fact]
        public void PointColors_ResolveSparseThemeFillWithoutAllocatingForUnboundedIndex() {
            using PowerPointDocument presentation = PowerPointDocument.Create();
            presentation.AddSlide().AddChartCm(OfficeChartType.Pie, CreateData(OfficeChartType.Pie, null), 1, 1, 20, 10);
            using var bytes = new MemoryStream();
            byte[] package = presentation.ToBytes();
            bytes.Write(package, 0, package.Length);
            bytes.Position = 0;
            OfficeColor expected;
            using (PresentationDocument native = PresentationDocument.Open(bytes, true)) {
                var part = native.PresentationPart!.SlideParts.Single().ChartParts.Single();
                C.PieChartSeries series = part.ChartSpace.Descendants<C.PieChartSeries>().Single();
                series.Append(new C.DataPoint(new C.Index { Val = 1U },
                    new C.ChartShapeProperties(new A.SolidFill(new A.SchemeColor { Val = A.SchemeColorValues.Accent1 }))));
                series.Append(new C.DataPoint(new C.Index { Val = uint.MaxValue },
                    new C.ChartShapeProperties(new A.SolidFill(new A.RgbColorModelHex { Val = "FF0000" }))));
                string rgb = native.PresentationPart.ThemePart!.Theme.ThemeElements!.ColorScheme!
                    .GetFirstChild<A.Accent1Color>()!.GetFirstChild<A.RgbColorModelHex>()!.Val!.Value!;
                expected = OfficeColor.FromRgb(Convert.ToByte(rgb.Substring(0, 2), 16),
                    Convert.ToByte(rgb.Substring(2, 2), 16), Convert.ToByte(rgb.Substring(4, 2), 16));
                part.ChartSpace.Save();
            }
            bytes.Position = 0;
            using PowerPointDocument reopened = PowerPointDocument.Load(bytes);
            AssertColors(Assert.Single(reopened.Slides.Single().Charts), new OfficeColor?[] { null, expected, null });
        }

        [Theory]
        [InlineData(OfficeChartType.Pie)]
        [InlineData(OfficeChartType.Doughnut)]
        public void ModernRadialChartProjectsDirectThemeColorStyle(OfficeChartType kind) {
            using PowerPointDocument presentation = PowerPointDocument.Create();
            PowerPointChart chart = presentation.AddSlide().AddChartCm(kind,
                new OfficeChartData(new[] { "A", "B" }, new[] {
                    new OfficeSeries("Values", new[] { 3d, 2d })
                }), 1, 1, 20, 10);
            Assert.True(chart.TryGetOfficeSnapshot(out OfficeChartSnapshot snapshot));
            string first = presentation.OpenXmlDocument.PresentationPart.ThemePart!.Theme.ThemeElements!
                .ColorScheme!.GetFirstChild<A.Accent1Color>()!.GetFirstChild<A.RgbColorModelHex>()!.Val!.Value!;
            string second = presentation.OpenXmlDocument.PresentationPart.ThemePart.Theme.ThemeElements
                .ColorScheme.GetFirstChild<A.Accent2Color>()!.GetFirstChild<A.RgbColorModelHex>()!.Val!.Value!;
            OfficeColor[] expected = new[] { first, second }.Select(OfficeColor.Parse).ToArray();
            Assert.Equal(expected, snapshot.Style.Palette.Take(2));
            Assert.Null(snapshot.Data.Series.Single().PointColors);
        }

        [Fact]
        public void RadialPaletteUsesChartLocalColorMapBeforeSlideMapping() {
            using PowerPointDocument presentation = PowerPointDocument.Create();
            PowerPointChart chart = presentation.AddSlide().AddChartCm(OfficeChartType.Pie,
                new OfficeChartData(new[] { "A", "B" }, new[] {
                    new OfficeSeries("Values", new[] { 3d, 2d })
                }), 1, 1, 20, 10);
            ChartPart part = presentation.Slides.Single().SlidePart.ChartParts.Single();
            part.ChartSpace!.AddChild(new C.ColorMapOverride {
                Background1 = A.ColorSchemeIndexValues.Light1,
                Text1 = A.ColorSchemeIndexValues.Dark1,
                Background2 = A.ColorSchemeIndexValues.Light2,
                Text2 = A.ColorSchemeIndexValues.Dark2,
                Accent1 = A.ColorSchemeIndexValues.Accent2,
                Accent2 = A.ColorSchemeIndexValues.Accent1,
                Accent3 = A.ColorSchemeIndexValues.Accent3,
                Accent4 = A.ColorSchemeIndexValues.Accent4,
                Accent5 = A.ColorSchemeIndexValues.Accent5,
                Accent6 = A.ColorSchemeIndexValues.Accent6,
                Hyperlink = A.ColorSchemeIndexValues.Hyperlink,
                FollowedHyperlink = A.ColorSchemeIndexValues.FollowedHyperlink
            }, true);
            chart.SetChartAreaStyle(fillColor: "ABCDEF");
            C.ShapeProperties area = part.ChartSpace.GetFirstChild<C.ShapeProperties>()!;
            area.GetFirstChild<A.SolidFill>()!.Remove();
            area.AddChild(new A.SolidFill(new A.SchemeColor { Val = A.SchemeColorValues.Accent1 }), true);

            Assert.True(chart.TryGetOfficeSnapshot(out OfficeChartSnapshot snapshot));
            A.ColorScheme scheme = presentation.OpenXmlDocument.PresentationPart.ThemePart!.Theme.ThemeElements!.ColorScheme!;
            OfficeColor expected = OfficeColor.Parse(scheme.GetFirstChild<A.Accent2Color>()!
                .GetFirstChild<A.RgbColorModelHex>()!.Val!.Value!);
            Assert.Equal(expected, snapshot.Style.Palette[0]);
            Assert.Equal(expected, snapshot.Style.BackgroundColor);
            Assert.Empty(presentation.ValidateDocument());
        }

        [Fact]
        public void MultiRingDoughnutCanUpdateDataWithoutProjectingItsPalette() {
            using PowerPointDocument presentation = PowerPointDocument.Create();
            var categories = new[] { "A", "B" };
            PowerPointChart chart = presentation.AddSlide().AddChartCm(OfficeChartType.Doughnut,
                new OfficeChartData(categories, new[] {
                    new OfficeSeries("Inner", new[] { 3d, 2d }),
                    new OfficeSeries("Outer", new[] { 4d, 1d })
                }), 1, 1, 20, 10);
            chart.UpdateData(new OfficeChartData(categories, new[] {
                new OfficeSeries("Inner", new[] { 5d, 6d }),
                new OfficeSeries("Outer", new[] { 7d, 8d })
            }));
            Assert.Empty(presentation.ValidateDocument());
        }

        [Fact]
        public void MultiRingDoughnutProjectsWhenEverySliceHasAnExplicitFill() {
            using PowerPointDocument presentation = PowerPointDocument.Create();
            OfficeColor?[] colors = { OfficeColor.Parse("#234567"), OfficeColor.Parse("#89ABCD") };
            PowerPointChart chart = presentation.AddSlide().AddChartCm(OfficeChartType.Doughnut,
                new OfficeChartData(new[] { "A", "B" }, new[] {
                    new OfficeSeries("Inner", new[] { 3d, 2d }, null, null, colors),
                    new OfficeSeries("Outer", new[] { 4d, 1d }, null, null, colors)
                }), 1, 1, 20, 10);
            Assert.True(chart.TryGetOfficeSnapshot(out OfficeChartSnapshot snapshot));
            Assert.Equal(2, snapshot.Data.Series.Count);
            Assert.All(snapshot.Data.Series, series => Assert.Equal(colors, series.PointColors));
        }

        [Theory]
        [InlineData(OfficeChartType.Pie)]
        [InlineData(OfficeChartType.Doughnut)]
        [InlineData(OfficeChartType.ColumnClustered)]
        public void PointColors_AppearInSlidePngAndRenderedPdf(OfficeChartType kind) {
            using PowerPointDocument authored = PowerPointDocument.Create();
            authored.SlideSize.SetSize(OfficeIMO.OfficeLength.Points(640),OfficeIMO.OfficeLength.Points(360));
            authored.AddSlide().AddChartPoints(kind, CreateData(kind, Colors), 20, 20, 600, 320);
            using var bytes = new MemoryStream(authored.ToBytes());
            using PowerPointDocument reopened = PowerPointDocument.Load(bytes);
            AssertColorPixels(reopened.Slides.Single().ToPng(), Colors[0]!.Value, Colors[2]!.Value);
            var options = new PowerPointPdfExportOptions();
            byte[] pdf = reopened.ToPdf(options).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true });
            Assert.DoesNotContain(options.Warnings, warning => warning.Code == "unsupported-chart");
            PdfCore.PdfPageRenderResult page = Assert.Single(PdfCore.PdfDocument.Load(pdf).Render.Pages(options:
                new PdfCore.PdfPageRenderOptions { Dpi = 72, Format = global::OfficeIMO.Drawing.OfficeImageExportFormat.Png, MaxPages = 1, ContinueOnError = false }));
            AssertColorPixels(page.Bytes!, Colors[0]!.Value, Colors[2]!.Value);
        }

        private static OfficeChartData CreateData(OfficeChartType kind, OfficeColor?[]? colors) =>
            new OfficeChartData(new[] { "Pass", "Unknown", "Fail" }, new[] {
                new OfficeSeries("Results", new double[] { 3, 2, 1 },
                    kind == OfficeChartType.Scatter ? new double[] { 1, 2, 3 } : null,
                    OfficeColor.FromRgb(90, 100, 110), colors)
            });

        private static void AssertColors(PowerPointChart chart, OfficeColor?[] expected) {
            Assert.True(chart.TryGetOfficeSnapshot(out OfficeChartSnapshot snapshot));
            Assert.Equal(expected, Assert.Single(snapshot.Data.Series).PointColors);
        }

        private static void AssertColorPixels(byte[] bytes, params OfficeColor[] expected) {
            Assert.True(OfficePngReader.TryDecode(bytes, out OfficeRasterImage? raster));
            foreach (OfficeColor color in expected) {
                int count = 0;
                for (int y = 0; y < raster!.Height; y++)
                    for (int x = 0; x < raster.Width; x++)
                        if (raster.GetPixel(x, y).Equals(color)) count++;
                Assert.True(count > 100, "Expected more than 100 pixels in authored point colour; found " + count);
            }
        }
    }
}
