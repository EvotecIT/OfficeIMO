using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.Html;
using OfficeIMO.PowerPoint.Pdf;
using Xunit;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using S = DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Tests {
    public class PowerPointSharedChartAuthoring {
        [Fact]
        public void SharedChartContract_AuthorsEveryKindAsValidNativeChart() {
            string output = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".pptx");
            string summaryPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".txt");
            OfficeChartKind[] kinds = Enum.GetValues(typeof(OfficeChartKind)).Cast<OfficeChartKind>().ToArray();
            try {
                using (PowerPointPresentation presentation = PowerPointPresentation.Create(output)) {
                    presentation.SlideSize.SetPreset(PowerPointSlideSizePreset.Screen16x9);
                    foreach (OfficeChartKind kind in kinds) {
                        PowerPointSlide slide = presentation.AddSlide();
                        OfficeChartData data = CreateData(kind);
                        PowerPointChart chart = slide.AddChartCm(kind, data, 1.5, 1.5, 22.4, 10.5,
                            new PowerPointChartAccessibilityOptions {
                                Name = kind + " Shared Chart",
                                AlternativeText = kind + " performance chart"
                            });
                        Assert.Contains("Data summary:", chart.AltText);
                        Assert.True(chart.TryGetOfficeSnapshot(out OfficeChartSnapshot snapshot));
                        Assert.Equal(kind, snapshot.ChartKind);
                        Assert.Equal(data.Categories, snapshot.Data.Categories);
                        if (kind == OfficeChartKind.ColumnClustered) {
                            chart.SaveDataSummary(summaryPath);
                        }
                    }

                    List<ValidationErrorInfo> errors = presentation.ValidateDocument();
                    Assert.True(errors.Count == 0, string.Join(Environment.NewLine,
                        errors.Select(error => error.Description)));
                    presentation.Save();
                }

                Assert.Contains("Chart kind: ColumnClustered", File.ReadAllText(summaryPath));
                using (PowerPointPresentation reopened = PowerPointPresentation.Open(output)) {
                    Assert.Equal(kinds.Length, reopened.Slides.Count);
                    for (int index = 0; index < kinds.Length; index++) {
                        PowerPointChart chart = Assert.Single(reopened.Slides[index].Charts);
                        Assert.True(chart.TryGetOfficeSnapshot(out OfficeChartSnapshot snapshot));
                        Assert.Equal(kinds[index], snapshot.ChartKind);
                    }
                }
                using (PresentationDocument document = PresentationDocument.Open(output, false)) {
                    List<ChartPart> chartParts = document.PresentationPart!.SlideParts
                        .SelectMany(part => part.ChartParts).ToList();
                    Assert.Equal(kinds.Length, chartParts.Count);
                    Assert.All(chartParts,
                        part => Assert.Single(part.GetPartsOfType<EmbeddedPackagePart>()));
                    EmbeddedPackagePart workbookPart = Assert.Single(document.PresentationPart.SlideParts.First()
                        .ChartParts.Single().GetPartsOfType<EmbeddedPackagePart>());
                    using SpreadsheetDocument workbook = SpreadsheetDocument.Open(workbookPart.GetStream(), false);
                    S.Cell actualValue = workbook.WorkbookPart!.WorksheetParts.Single().Worksheet
                        .Descendants<S.Cell>().Single(cell => cell.CellReference?.Value == "B2");
                    Assert.Equal("28", actualValue.CellValue?.Text);
                }
            } finally {
                if (File.Exists(output)) File.Delete(output);
                if (File.Exists(summaryPath)) File.Delete(summaryPath);
            }
        }

        [Fact]
        public void SharedChartContract_AuthorsComboChartWithSecondaryAxisAndSeriesKinds() {
            string output = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".pptx");
            OfficeChartData data = CreateComboData();
            try {
                using (PowerPointPresentation presentation = PowerPointPresentation.Create(output)) {
                    PowerPointSlide slide = presentation.AddSlide();
                    PowerPointChart chart = slide.AddChartCm(OfficeChartKind.ColumnClustered, data,
                        1.5, 1.5, 22, 10, new PowerPointChartAccessibilityOptions {
                            AlternativeText = "Revenue columns with margin line on a secondary axis"
                        });
                    Assert.True(chart.TryGetOfficeSnapshot(out OfficeChartSnapshot snapshot));
                    Assert.Equal(2, snapshot.Data.Series.Count);
                    Assert.Equal(OfficeChartKind.ColumnClustered, snapshot.Data.Series[0].RenderKind);
                    Assert.Equal(OfficeChartKind.Line, snapshot.Data.Series[1].RenderKind);
                    Assert.Equal(OfficeChartAxisGroup.Primary, snapshot.Data.Series[0].AxisGroup);
                    Assert.Equal(OfficeChartAxisGroup.Secondary, snapshot.Data.Series[1].AxisGroup);
                    Assert.Contains("Revenue", chart.CreateDataSummary());
                    List<ValidationErrorInfo> errors = presentation.ValidateDocument();
                    Assert.True(errors.Count == 0, string.Join(Environment.NewLine,
                        errors.Select(error => error.Description)));
                    presentation.Save();
                }

                using (PresentationDocument document = PresentationDocument.Open(output, false)) {
                    ChartPart chartPart = Assert.Single(document.PresentationPart!.SlideParts
                        .SelectMany(part => part.ChartParts));
                    C.PlotArea plotArea = chartPart.ChartSpace!.GetFirstChild<C.Chart>()!
                        .GetFirstChild<C.PlotArea>()!;
                    Assert.Single(plotArea.Elements<C.BarChart>());
                    Assert.Single(plotArea.Elements<C.LineChart>());
                    Assert.Equal(2, plotArea.Elements<C.CategoryAxis>().Count());
                    Assert.Equal(2, plotArea.Elements<C.ValueAxis>().Count());
                    Assert.Contains(plotArea.Elements<C.ValueAxis>(), axis =>
                        axis.AxisPosition?.Val?.Value == AxisPositionValues.Right);
                }
            } finally {
                if (File.Exists(output)) File.Delete(output);
            }
        }

        [Fact]
        public void SharedChartContract_PreservesHiddenLegendSeriesInValidNativeChartAndSnapshot() {
            var data = new OfficeChartData(new[] { "Q1", "Q2" }, new[] {
                new OfficeChartSeries("Visible", new[] { 12D, 18D }, null, null, null,
                    showMarkers: false, showInLegend: true),
                new OfficeChartSeries("Threshold", new[] { 15D, 15D }, null, null, null,
                    showMarkers: false, showInLegend: false)
            });
            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream, autoSave: false);
            PowerPointSlide slide = presentation.Slides[0];
            PowerPointChart chart = slide.AddChart(OfficeChartKind.ColumnClustered, data);

            C.Legend legend = slide.SlidePart.ChartParts.Single().ChartSpace!.GetFirstChild<C.Chart>()!
                .GetFirstChild<C.Legend>()!;
            Assert.IsType<C.LegendPosition>(legend.ChildElements[0]);
            C.LegendEntry entry = Assert.IsType<C.LegendEntry>(legend.ChildElements[1]);
            Assert.Equal(1U, entry.Index!.Val!.Value);
            Assert.True(entry.GetFirstChild<C.Delete>()!.Val!.Value);
            Assert.True(chart.TryGetOfficeSnapshot(out OfficeChartSnapshot snapshot));
            Assert.True(snapshot.Data.Series[0].ShowInLegend);
            Assert.False(snapshot.Data.Series[1].ShowInLegend);
            Assert.Empty(presentation.ValidateDocument());
        }

        [Fact]
        public void SharedChartSnapshot_ClassifiesTopValueAxisAsSecondaryForHorizontalBars() {
            string output = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".pptx");
            var data = new OfficeChartData(new[] { "North", "South" }, new[] {
                new OfficeChartSeries("Primary", new[] { 42D, 55D }, null, null, null,
                    showMarkers: false, renderKind: OfficeChartKind.ColumnClustered),
                new OfficeChartSeries("Secondary", new[] { 4.2D, 5.5D }, null, null, null,
                    showMarkers: false, renderKind: OfficeChartKind.ColumnClustered,
                    axisGroup: OfficeChartAxisGroup.Secondary)
            });
            try {
                using (PowerPointPresentation presentation = PowerPointPresentation.Create(output)) {
                    PowerPointSlide slide = presentation.Slides[0];
                    slide.AddChart(OfficeChartKind.ColumnClustered, data);
                    ChartPart chartPart = slide.SlidePart.ChartParts.Single();
                    C.PlotArea plotArea = chartPart.ChartSpace!.GetFirstChild<C.Chart>()!
                        .GetFirstChild<C.PlotArea>()!;
                    foreach (C.BarChart barChart in plotArea.Elements<C.BarChart>()) {
                        barChart.BarDirection!.Val = C.BarDirectionValues.Bar;
                    }
                    C.BarChart secondaryChart = plotArea.Elements<C.BarChart>().Last();
                    uint[] secondaryAxisIds = secondaryChart.Elements<C.AxisId>()
                        .Select(axis => axis.Val!.Value).ToArray();
                    plotArea.Elements<C.CategoryAxis>().Single(axis =>
                        axis.AxisId?.Val?.Value == secondaryAxisIds[0]).AxisPosition!.Val =
                        C.AxisPositionValues.Right;
                    plotArea.Elements<C.ValueAxis>().Single(axis =>
                        axis.AxisId?.Val?.Value == secondaryAxisIds[1]).AxisPosition!.Val =
                        C.AxisPositionValues.Top;
                    chartPart.ChartSpace.Save();
                    presentation.Save();
                }

                using PowerPointPresentation reopened = PowerPointPresentation.OpenRead(output);
                PowerPointChart chart = Assert.Single(reopened.Slides[0].Charts);
                Assert.True(chart.TryGetOfficeSnapshot(out OfficeChartSnapshot snapshot));
                Assert.Equal(OfficeChartAxisGroup.Primary, snapshot.Data.Series[0].AxisGroup);
                Assert.Equal(OfficeChartAxisGroup.Secondary, snapshot.Data.Series[1].AxisGroup);
            } finally {
                if (File.Exists(output)) File.Delete(output);
            }
        }

        [Fact]
        public void SharedScatterChartUsesExplicitXValuesWithNonNumericLabels() {
            var data = new OfficeChartData(new[] { "Discovery", "Delivery" }, new[] {
                new OfficeChartSeries("Actual", new[] { 10D, 14D }, new[] { 1.25D, 2.75D })
            });
            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream, autoSave: false);

            PowerPointChart chart = presentation.Slides[0].AddChart(
                OfficeChartKind.Scatter, data);

            Assert.True(chart.TryGetOfficeSnapshot(out OfficeChartSnapshot snapshot));
            Assert.Equal(new[] { 1.25D, 2.75D }, Assert.Single(snapshot.Data.Series).XValues);
        }

        [Fact]
        public void GeneratedChartSummaryCarriesAccessibilityMarker() {
            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream, autoSave: false);
            PowerPointChart chart = presentation.Slides[0].AddChart(
                OfficeChartKind.ColumnClustered, CreateData(OfficeChartKind.ColumnClustered),
                accessibility: new PowerPointChartAccessibilityOptions());

            Assert.StartsWith("Data summary:", chart.AltText, StringComparison.Ordinal);
            PowerPointAccessibilityReport report = presentation.InspectAccessibility();
            Assert.DoesNotContain(report.Findings,
                finding => finding.Code == "Accessibility.ChartColorOnlyMeaning");
        }

        [Fact]
        public void ScatterDataSummaryIncludesEveryExplicitPoint() {
            var data = new OfficeChartData(new[] { "First" }, new[] {
                new OfficeChartSeries("Actual", new[] { 10D, 20D, 30D }, new[] { 1D, 2D, 3D })
            });

            string summary = PowerPointChart.CreateDataSummary(OfficeChartKind.Scatter, data);

            Assert.Contains("Series\tX\tY", summary, StringComparison.Ordinal);
            Assert.Contains("Actual\t1\t10", summary, StringComparison.Ordinal);
            Assert.Contains("Actual\t2\t20", summary, StringComparison.Ordinal);
            Assert.Contains("Actual\t3\t30", summary, StringComparison.Ordinal);
        }

        [Fact]
        public void SharedChartContract_RejectsUnsupportedMixedAxesBeforeMutatingSlide() {
            var data = new OfficeChartData(new[] { "A", "B" }, new[] {
                new OfficeChartSeries("Bars", new[] { 1D, 2D }, null, null, null, showMarkers: false,
                    renderKind: OfficeChartKind.BarClustered),
                new OfficeChartSeries("Line", new[] { 2D, 3D }, null, null, null, showMarkers: true,
                    renderKind: OfficeChartKind.Line, axisGroup: OfficeChartAxisGroup.Secondary)
            });
            using PowerPointPresentation presentation = PowerPointPresentation.Create(Stream.Null);
            PowerPointSlide slide = presentation.AddSlide();

            Assert.Throws<NotSupportedException>(() =>
                slide.AddChart(OfficeChartKind.BarClustered, data));
            Assert.Empty(slide.Charts);

            var mismatchedScatter = new OfficeChartData(new[] { "1", "2" }, new[] {
                new OfficeChartSeries("Observed", new[] { 2D, 4D }, new[] { 1D, 2D }, null, null,
                    showMarkers: true, renderKind: OfficeChartKind.Scatter)
            });
            Assert.Throws<NotSupportedException>(() =>
                slide.AddChart(OfficeChartKind.ColumnClustered, mismatchedScatter));
            Assert.Empty(slide.Charts);
        }

        [Fact]
        public void SharedChartContract_DrivesPngSvgHtmlAndPdfFromAuthoredSnapshots() {
            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
            presentation.SlideSize.SetSizePoints(360, 220);
            foreach (OfficeChartKind kind in new[] {
                         OfficeChartKind.AreaStacked,
                         OfficeChartKind.Radar,
                         OfficeChartKind.BarStacked100
                     }) {
                PowerPointSlide slide = presentation.AddSlide();
                slide.AddChartPoints(kind, CreateData(kind), 30, 20, 300, 170,
                    new PowerPointChartAccessibilityOptions { AlternativeText = kind + " export proof" });
                OfficeImageExportResult png = slide.ExportImage(OfficeImageExportFormat.Png);
                OfficeImageExportResult svg = slide.ExportImage(OfficeImageExportFormat.Svg);
                Assert.True(png.Bytes.Length > 100);
                Assert.Contains("<svg", System.Text.Encoding.UTF8.GetString(svg.Bytes),
                    StringComparison.Ordinal);
                Assert.DoesNotContain(png.Diagnostics, diagnostic => diagnostic.Severity == OfficeImageExportDiagnosticSeverity.Error);
                Assert.DoesNotContain(svg.Diagnostics, diagnostic => diagnostic.Severity == OfficeImageExportDiagnosticSeverity.Error);
            }
            PowerPointSlide comboSlide = presentation.AddSlide();
            comboSlide.AddChartPoints(OfficeChartKind.ColumnClustered, CreateComboData(), 30, 20, 300, 170,
                new PowerPointChartAccessibilityOptions { AlternativeText = "Combo export proof" });
            OfficeImageExportResult comboPng = comboSlide.ExportImage(OfficeImageExportFormat.Png);
            OfficeImageExportResult comboSvg = comboSlide.ExportImage(OfficeImageExportFormat.Svg);
            Assert.DoesNotContain(comboPng.Diagnostics,
                diagnostic => diagnostic.Severity == OfficeImageExportDiagnosticSeverity.Error);
            Assert.DoesNotContain(comboSvg.Diagnostics,
                diagnostic => diagnostic.Severity == OfficeImageExportDiagnosticSeverity.Error);

            string html = presentation.ToHtml(new PowerPointHtmlSaveOptions {
                Profile = OfficeHtmlConversionProfile.PowerPointSemanticSlides
            });
            byte[] pdf = presentation.SaveAsPdf();

            Assert.Contains("StackedArea", html, StringComparison.Ordinal);
            Assert.Contains("Radar", html, StringComparison.Ordinal);
            Assert.True(pdf.Length > 500);
        }

        private static OfficeChartData CreateData(OfficeChartKind kind) {
            if (kind == OfficeChartKind.Scatter) {
                return new OfficeChartData(new[] { "1", "2", "3", "4" }, new[] {
                    new OfficeChartSeries("Observed", new[] { 2D, 4D, 3D, 5D },
                        new[] { 1D, 2D, 3D, 4D }, OfficeColor.Parse("#0B7FAB"))
                });
            }
            if (kind == OfficeChartKind.Pie || kind == OfficeChartKind.Doughnut) {
                return new OfficeChartData(new[] { "Services", "Licenses", "Support" }, new[] {
                    new OfficeChartSeries("Share", new[] { 55D, 30D, 15D }, null,
                        OfficeColor.Parse("#0B7FAB"), new OfficeColor?[] {
                            OfficeColor.Parse("#0B7FAB"), OfficeColor.Parse("#4CAF50"), OfficeColor.Parse("#E85D04")
                        })
                });
            }
            return new OfficeChartData(new[] { "Q1", "Q2", "Q3", "Q4" }, new[] {
                new OfficeChartSeries("Actual", new[] { 28D, 43D, 61D, 72D }, null,
                    OfficeColor.Parse("#0B7FAB"), null, showMarkers: true,
                    markerSize: 7, strokeWidth: 1.8),
                new OfficeChartSeries("Target", new[] { 35D, 50D, 65D, 80D }, null,
                    OfficeColor.Parse("#4CAF50"), null, showMarkers: true,
                    markerSize: 7, strokeWidth: 1.8, strokeDashStyle: OfficeStrokeDashStyle.Dash)
            });
        }

        private static OfficeChartData CreateComboData() => new(
            new[] { "Q1", "Q2", "Q3", "Q4" }, new[] {
                new OfficeChartSeries("Revenue", new[] { 120D, 145D, 172D, 190D }, null,
                    OfficeColor.Parse("#0B7FAB"), null, showMarkers: false,
                    renderKind: OfficeChartKind.ColumnClustered),
                new OfficeChartSeries("Margin", new[] { 22D, 26D, 31D, 35D }, null,
                    OfficeColor.Parse("#E85D04"), null, showMarkers: true,
                    markerSize: 8, markerShape: OfficeChartMarkerShape.Circle, strokeWidth: 2,
                    renderKind: OfficeChartKind.Line, axisGroup: OfficeChartAxisGroup.Secondary)
            });
    }
}
