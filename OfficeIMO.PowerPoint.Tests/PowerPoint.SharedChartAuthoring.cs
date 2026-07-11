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
using A = DocumentFormat.OpenXml.Drawing;
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
                using (PowerPointPresentation reopened = PowerPointPresentation.Open(output, PowerPointOpenMode.ReadOnly)) {
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
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream, new PowerPointStreamCreateOptions { AutoSave = false });
            PowerPointSlide slide = presentation.AddSlide();
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
        public void SharedChartUpdate_RebuildsRadarDataAndLegendMetadataWithoutLosingTitle() {
            var initial = new OfficeChartData(new[] { "Q1", "Q2", "Q3" }, new[] {
                new OfficeChartSeries("Actual", new[] { 10D, 12D, 14D }),
                new OfficeChartSeries("Target", new[] { 11D, 13D, 15D })
            });
            var updated = new OfficeChartData(new[] { "Q1", "Q2", "Q3" }, new[] {
                new OfficeChartSeries("Actual", new[] { 20D, 22D, 24D }, null, null, null,
                    showMarkers: true, showInLegend: true),
                new OfficeChartSeries("Target", new[] { 21D, 23D, 25D }, null, null, null,
                    showMarkers: true, showInLegend: false)
            });
            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream,
                new PowerPointStreamCreateOptions { AutoSave = false });
            PowerPointSlide slide = presentation.AddSlide();
            PowerPointChart chart = slide.AddChart(OfficeChartKind.Radar, initial).SetTitle("Radar update");

            chart.UpdateData(updated);

            Assert.True(chart.TryGetOfficeSnapshot(out OfficeChartSnapshot snapshot));
            Assert.Equal(OfficeChartKind.Radar, snapshot.ChartKind);
            Assert.Equal("Radar update", snapshot.Title);
            Assert.Equal(new[] { 20D, 22D, 24D }, snapshot.Data.Series[0].Values);
            Assert.False(snapshot.Data.Series[1].ShowInLegend);
            C.PlotArea plotArea = slide.SlidePart.ChartParts.Single().ChartSpace!
                .GetFirstChild<C.Chart>()!.GetFirstChild<C.PlotArea>()!;
            Assert.Equal(2, Assert.Single(plotArea.Elements<C.RadarChart>())
                .Elements<C.RadarChartSeries>().Count());
            Assert.Empty(presentation.ValidateDocument());
        }

        [Fact]
        public void SharedChartUpdate_RebuildsEveryComboLayerAndSecondaryAxisSeries() {
            var updated = new OfficeChartData(new[] { "Q1", "Q2", "Q3", "Q4" }, new[] {
                new OfficeChartSeries("Revenue", new[] { 210D, 230D, 250D, 280D }, null,
                    OfficeColor.Parse("#0B7FAB"), null, showMarkers: false,
                    renderKind: OfficeChartKind.ColumnClustered),
                new OfficeChartSeries("Margin", new[] { 31D, 34D, 37D, 40D }, null,
                    OfficeColor.Parse("#E85D04"), null, showMarkers: true, showInLegend: false,
                    renderKind: OfficeChartKind.Line, axisGroup: OfficeChartAxisGroup.Secondary)
            });
            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream,
                new PowerPointStreamCreateOptions { AutoSave = false });
            PowerPointSlide slide = presentation.AddSlide();
            PowerPointChart chart = slide.AddChart(OfficeChartKind.ColumnClustered, CreateComboData());

            chart.UpdateData(updated);

            Assert.True(chart.TryGetOfficeSnapshot(out OfficeChartSnapshot snapshot));
            Assert.Equal(new[] { 210D, 230D, 250D, 280D }, snapshot.Data.Series[0].Values);
            Assert.Equal(new[] { 31D, 34D, 37D, 40D }, snapshot.Data.Series[1].Values);
            Assert.Equal(OfficeChartKind.Line, snapshot.Data.Series[1].RenderKind);
            Assert.Equal(OfficeChartAxisGroup.Secondary, snapshot.Data.Series[1].AxisGroup);
            Assert.False(snapshot.Data.Series[1].ShowInLegend);
            C.PlotArea plotArea = slide.SlidePart.ChartParts.Single().ChartSpace!
                .GetFirstChild<C.Chart>()!.GetFirstChild<C.PlotArea>()!;
            Assert.Single(plotArea.Elements<C.BarChart>());
            Assert.Single(plotArea.Elements<C.LineChart>());
            Assert.Empty(presentation.ValidateDocument());
        }

        [Fact]
        public void SharedChartUpdate_PreservesAxisCustomization() {
            OfficeChartData initial = CreateData(OfficeChartKind.ColumnClustered);
            var updated = new OfficeChartData(initial.Categories, new[] {
                new OfficeChartSeries("Revenue", new[] { 40D, 45D, 52D, 61D }),
                new OfficeChartSeries("Margin", new[] { 20D, 23D, 27D, 31D })
            });
            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream,
                new PowerPointStreamCreateOptions { AutoSave = false });
            PowerPointSlide slide = presentation.AddSlide();
            PowerPointChart chart = slide.AddChart(OfficeChartKind.ColumnClustered, initial)
                .SetCategoryAxisTitle("Quarter")
                .SetValueAxisTitle("Revenue")
                .SetValueAxisNumberFormat("$#,##0")
                .SetValueAxisGridlines(showMajor: true, showMinor: true, lineColor: "D9E2F3");

            chart.UpdateData(updated);

            C.PlotArea plotArea = slide.SlidePart.ChartParts.Single().ChartSpace!
                .GetFirstChild<C.Chart>()!.GetFirstChild<C.PlotArea>()!;
            C.CategoryAxis categoryAxis = Assert.Single(plotArea.Elements<C.CategoryAxis>());
            C.ValueAxis valueAxis = Assert.Single(plotArea.Elements<C.ValueAxis>());
            Assert.Equal("Quarter", string.Concat(categoryAxis.Descendants<A.Text>().Select(text => text.Text)));
            Assert.Equal("Revenue", string.Concat(valueAxis.Descendants<A.Text>().Select(text => text.Text)));
            Assert.Equal("$#,##0", valueAxis.GetFirstChild<C.NumberingFormat>()!.FormatCode!.Value);
            Assert.NotNull(valueAxis.GetFirstChild<C.MajorGridlines>());
            Assert.NotNull(valueAxis.GetFirstChild<C.MinorGridlines>());
            Assert.Empty(presentation.ValidateDocument());
        }

        [Fact]
        public void SharedChartUpdate_PreservesHiddenLegend() {
            OfficeChartData initial = CreateData(OfficeChartKind.ColumnClustered);
            var updated = new OfficeChartData(initial.Categories, new[] {
                new OfficeChartSeries("Revenue", new[] { 40D, 45D, 52D, 61D })
            });
            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream,
                new PowerPointStreamCreateOptions { AutoSave = false });
            PowerPointSlide slide = presentation.AddSlide();
            PowerPointChart chart = slide.AddChart(OfficeChartKind.ColumnClustered, initial)
                .HideLegend();

            chart.UpdateData(updated);

            C.Chart nativeChart = slide.SlidePart.ChartParts.Single().ChartSpace!
                .GetFirstChild<C.Chart>()!;
            Assert.Null(nativeChart.GetFirstChild<C.Legend>());
            Assert.Empty(presentation.ValidateDocument());
        }

        [Fact]
        public void SharedChartUpdate_PreservesCompatibleChartAndSeriesFormatting() {
            OfficeChartData initial = CreateData(OfficeChartKind.Line);
            var updated = new OfficeChartData(initial.Categories, new[] {
                new OfficeChartSeries("Actual", new[] { 31D, 47D, 64D, 76D }),
                new OfficeChartSeries("Target", new[] { 36D, 52D, 68D, 83D })
            });
            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream,
                new PowerPointStreamCreateOptions { AutoSave = false });
            PowerPointSlide slide = presentation.AddSlide();
            PowerPointChart chart = slide.AddChart(OfficeChartKind.Line, initial)
                .SetDataLabels(showValue: true)
                .SetSeriesDataLabels(0, showValue: true, numberFormat: "0.0")
                .SetSeriesTrendline(0, C.TrendlineValues.Linear, displayRSquared: true);

            chart.UpdateData(updated);

            C.LineChart lineChart = Assert.Single(slide.SlidePart.ChartParts.Single().ChartSpace!
                .GetFirstChild<C.Chart>()!.GetFirstChild<C.PlotArea>()!.Elements<C.LineChart>());
            Assert.True(lineChart.GetFirstChild<C.DataLabels>()!
                .GetFirstChild<C.ShowValue>()!.Val!.Value);
            C.LineChartSeries firstSeries = lineChart.Elements<C.LineChartSeries>().First();
            Assert.Equal("0.0", firstSeries.GetFirstChild<C.DataLabels>()!
                .GetFirstChild<C.NumberingFormat>()!.FormatCode!.Value);
            Assert.NotNull(firstSeries.GetFirstChild<C.Trendline>());
            Assert.Empty(presentation.ValidateDocument());
        }

        [Fact]
        public void SharedChartUpdate_RegeneratesAxesWhenOrientationChanges() {
            OfficeChartData initial = CreateData(OfficeChartKind.ColumnClustered);
            var updated = new OfficeChartData(initial.Categories, new[] {
                new OfficeChartSeries("Actual", new[] { 31D, 47D, 64D, 76D }, null, null, null,
                    showMarkers: false, renderKind: OfficeChartKind.BarClustered),
                new OfficeChartSeries("Target", new[] { 36D, 52D, 68D, 83D }, null, null, null,
                    showMarkers: false, renderKind: OfficeChartKind.BarClustered)
            });
            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream,
                new PowerPointStreamCreateOptions { AutoSave = false });
            PowerPointSlide slide = presentation.AddSlide();
            PowerPointChart chart = slide.AddChart(OfficeChartKind.ColumnClustered, initial)
                .SetCategoryAxisTitle("Quarter")
                .SetValueAxisTitle("Revenue");

            chart.UpdateData(updated);

            C.PlotArea plotArea = slide.SlidePart.ChartParts.Single().ChartSpace!
                .GetFirstChild<C.Chart>()!.GetFirstChild<C.PlotArea>()!;
            Assert.Equal(C.BarDirectionValues.Bar,
                Assert.Single(plotArea.Elements<C.BarChart>()).BarDirection!.Val!.Value);
            Assert.Equal(C.AxisPositionValues.Left,
                Assert.Single(plotArea.Elements<C.CategoryAxis>()).AxisPosition!.Val!.Value);
            Assert.Equal(C.AxisPositionValues.Bottom,
                Assert.Single(plotArea.Elements<C.ValueAxis>()).AxisPosition!.Val!.Value);
            Assert.Empty(presentation.ValidateDocument());
        }

        [Fact]
        public void SharedChartUpdate_PreservesImportedDateAxisSemantics() {
            var initial = new OfficeChartData(new[] { "2026-01-01", "2026-02-01", "2026-03-01" }, new[] {
                new OfficeChartSeries("Revenue", new[] { 12D, 18D, 25D })
            });
            var updated = new OfficeChartData(initial.Categories, new[] {
                new OfficeChartSeries("Revenue", new[] { 15D, 22D, 31D })
            });
            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream,
                new PowerPointStreamCreateOptions { AutoSave = false });
            PowerPointSlide slide = presentation.AddSlide();
            PowerPointChart chart = slide.AddChart(OfficeChartKind.Line, initial)
                .SetCategoryAxisTitle("Month");
            ChartPart chartPart = slide.SlidePart.ChartParts.Single();
            C.PlotArea plotArea = chartPart.ChartSpace!.GetFirstChild<C.Chart>()!
                .GetFirstChild<C.PlotArea>()!;
            C.DateAxis importedDateAxis = ReplaceWithDateAxis(plotArea,
                Assert.Single(plotArea.Elements<C.CategoryAxis>()));
            importedDateAxis.NumberingFormat = new C.NumberingFormat {
                FormatCode = "mmm-yy",
                SourceLinked = false
            };
            importedDateAxis.AddChild(new C.BaseTimeUnit { Val = C.TimeUnitValues.Days }, true);
            importedDateAxis.AddChild(new C.MajorUnit { Val = 2D }, true);
            importedDateAxis.AddChild(new C.MajorTimeUnit { Val = C.TimeUnitValues.Months }, true);
            chartPart.ChartSpace.Save();

            chart.UpdateData(updated);

            plotArea = chartPart.ChartSpace!.GetFirstChild<C.Chart>()!.GetFirstChild<C.PlotArea>()!;
            Assert.Empty(plotArea.Elements<C.CategoryAxis>());
            C.DateAxis dateAxis = Assert.Single(plotArea.Elements<C.DateAxis>());
            Assert.Equal("Month", string.Concat(dateAxis.Descendants<A.Text>().Select(text => text.Text)));
            Assert.Equal("mmm-yy", dateAxis.NumberingFormat!.FormatCode!.Value);
            Assert.Equal(C.TimeUnitValues.Days, dateAxis.GetFirstChild<C.BaseTimeUnit>()!.Val!.Value);
            Assert.Equal(2D, dateAxis.GetFirstChild<C.MajorUnit>()!.Val!.Value);
            Assert.Equal(C.TimeUnitValues.Months, dateAxis.GetFirstChild<C.MajorTimeUnit>()!.Val!.Value);
            Assert.Empty(presentation.ValidateDocument());
        }

        [Fact]
        public void SharedChartUpdate_RefreshesEveryImportedScatterLayer() {
            var initial = new OfficeChartData(new[] { "1", "2", "3" }, new[] {
                new OfficeChartSeries("Actual", new[] { 10D, 12D, 14D }, new[] { 1D, 2D, 3D }),
                new OfficeChartSeries("Forecast", new[] { 11D, 13D, 15D }, new[] { 1D, 2D, 3D })
            });
            var updated = new OfficeChartData(initial.Categories, new[] {
                new OfficeChartSeries("Actual", new[] { 20D, 24D, 28D }, new[] { 1D, 2D, 3D }),
                new OfficeChartSeries("Forecast", new[] { 21D, 25D, 29D }, new[] { 1D, 2D, 3D })
            });
            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream,
                new PowerPointStreamCreateOptions { AutoSave = false });
            PowerPointSlide slide = presentation.AddSlide();
            PowerPointChart chart = slide.AddChart(OfficeChartKind.Scatter, initial);
            ChartPart chartPart = slide.SlidePart.ChartParts.Single();
            C.PlotArea plotArea = chartPart.ChartSpace!.GetFirstChild<C.Chart>()!
                .GetFirstChild<C.PlotArea>()!;
            C.ScatterChart firstLayer = Assert.Single(plotArea.Elements<C.ScatterChart>());
            var secondLayer = (C.ScatterChart)firstLayer.CloneNode(true);
            firstLayer.Elements<C.ScatterChartSeries>().Last().Remove();
            secondLayer.Elements<C.ScatterChartSeries>().First().Remove();
            plotArea.InsertAfter(secondLayer, firstLayer);
            chartPart.ChartSpace.Save();

            chart.UpdateData(updated);

            C.ScatterChart[] layers = plotArea.Elements<C.ScatterChart>().ToArray();
            Assert.Equal(2, layers.Length);
            Assert.Equal(0U, Assert.Single(layers[0].Elements<C.ScatterChartSeries>()).Index!.Val!.Value);
            Assert.Equal(1U, Assert.Single(layers[1].Elements<C.ScatterChartSeries>()).Index!.Val!.Value);
            Assert.True(chart.TryGetOfficeSnapshot(out OfficeChartSnapshot snapshot));
            Assert.Equal(2, snapshot.Data.Series.Count);
            Assert.Equal(new[] { 20D, 24D, 28D }, snapshot.Data.Series[0].Values);
            Assert.Equal(new[] { 21D, 25D, 29D }, snapshot.Data.Series[1].Values);
            Assert.Empty(presentation.ValidateDocument());
        }

        [Fact]
        public void SharedChartUpdate_RejectsUnsupportedNativeChartWithoutMutation() {
            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream,
                new PowerPointStreamCreateOptions { AutoSave = false });
            PowerPointSlide slide = presentation.AddSlide();
            PowerPointChart chart = slide.AddChart(OfficeChartKind.ColumnClustered,
                CreateData(OfficeChartKind.ColumnClustered));
            C.PlotArea plotArea = slide.SlidePart.ChartParts.Single().ChartSpace!
                .GetFirstChild<C.Chart>()!.GetFirstChild<C.PlotArea>()!;
            plotArea.ReplaceChild(new C.BubbleChart(), Assert.Single(plotArea.Elements<C.BarChart>()));
            string originalPlotArea = plotArea.OuterXml;

            Assert.Throws<NotSupportedException>(() => chart.UpdateData(
                CreateData(OfficeChartKind.ColumnClustered)));

            Assert.Equal(originalPlotArea, plotArea.OuterXml);
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
                    PowerPointSlide slide = presentation.AddSlide();
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

                using PowerPointPresentation reopened = PowerPointPresentation.Open(output, PowerPointOpenMode.ReadOnly);
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
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream, new PowerPointStreamCreateOptions { AutoSave = false });

            PowerPointChart chart = presentation.AddSlide().AddChart(
                OfficeChartKind.Scatter, data);

            Assert.True(chart.TryGetOfficeSnapshot(out OfficeChartSnapshot snapshot));
            Assert.Equal(new[] { 1.25D, 2.75D }, Assert.Single(snapshot.Data.Series).XValues);
        }

        [Fact]
        public void GeneratedChartSummaryCarriesAccessibilityMarker() {
            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream, new PowerPointStreamCreateOptions { AutoSave = false });
            PowerPointChart chart = presentation.AddSlide().AddChart(
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

        private static C.DateAxis ReplaceWithDateAxis(C.PlotArea plotArea, C.CategoryAxis categoryAxis) {
            var dateAxis = new C.DateAxis();
            foreach (var child in categoryAxis.ChildElements) {
                if (child is C.AxisId or C.Scaling or C.Delete or C.AxisPosition or
                    C.MajorGridlines or C.MinorGridlines or C.Title or C.NumberingFormat or
                    C.MajorTickMark or C.MinorTickMark or C.TickLabelPosition or
                    C.ChartShapeProperties or C.TextProperties or C.CrossingAxis or
                    C.Crosses or C.CrossesAt or C.AutoLabeled or C.LabelOffset) {
                    dateAxis.AddChild(child.CloneNode(true), true);
                }
            }
            plotArea.ReplaceChild(dateAxis, categoryAxis);
            return dateAxis;
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
