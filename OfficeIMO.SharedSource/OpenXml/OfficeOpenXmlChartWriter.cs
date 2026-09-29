using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Drawing;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.OpenXml.Internal {
    internal static partial class OfficeOpenXmlChartWriter {
        private const int SpreadsheetMaximumColumns = 16_384;
        private const int SpreadsheetMaximumRows = 1_048_576;
        private const int BubbleWorkbookColumnsPerSeries = 3;
        internal const int MaximumSharedChartPoints = 100_000;

        private sealed class SharedSeriesDescriptor {
            internal SharedSeriesDescriptor(int index, OfficeChartSeries series, OfficeChartKind kind) {
                Index = index;
                Series = series;
                Kind = kind;
            }

            internal int Index { get; }
            internal OfficeChartSeries Series { get; }
            internal OfficeChartKind Kind { get; }
            internal OfficeChartAxisGroup AxisGroup => Series.AxisGroup;
        }

        internal static void PopulateSharedChart(ChartPart chartPart, string embeddedRelId, OfficeChartData data,
            OfficeChartKind defaultKind) {
            if (chartPart == null) throw new ArgumentNullException(nameof(chartPart));
            ValidateSharedChartData(data, defaultKind);

            C.ChartSpace chartSpace = new();
            chartSpace.AddNamespaceDeclaration("c", ChartNamespace);
            chartSpace.AddNamespaceDeclaration("a", DrawingNamespace);
            chartSpace.AddNamespaceDeclaration("r", RelationshipNamespace);
            chartSpace.Append(new C.Date1904 { Val = false });
            chartSpace.Append(new C.EditingLanguage { Val = "en-US" });
            chartSpace.Append(new C.RoundedCorners { Val = false });

            C.Chart chart = new(new C.AutoTitleDeleted { Val = false });
            C.PlotArea plotArea = new(new C.Layout());
            AppendSharedChartContent(plotArea, data, defaultKind);
            chart.Append(plotArea);
            chart.Append(CreateSharedLegend(data, defaultKind));
            chart.Append(new C.PlotVisibleOnly { Val = true });
            chart.Append(new C.DisplayBlanksAs { Val = C.DisplayBlanksAsValues.Gap });
            chart.Append(new C.ShowDataLabelsOverMaximum { Val = false });
            chartSpace.Append(chart);
            if (!string.IsNullOrWhiteSpace(embeddedRelId)) {
                chartSpace.Append(new C.ExternalData {
                    Id = embeddedRelId,
                    AutoUpdate = new C.AutoUpdate { Val = false }
                });
            }
            chartPart.ChartSpace = chartSpace;
            ApplySharedChartSeriesStyle(chartPart, data, defaultKind);
        }

        internal static void UpdateSharedChartData(ChartPart chartPart, OfficeChartData data,
            OfficeChartKind defaultKind) {
            if (chartPart == null) throw new ArgumentNullException(nameof(chartPart));
            ValidateSharedChartData(data, defaultKind);

            C.ChartSpace chartSpace = chartPart.ChartSpace ??
                throw new InvalidOperationException("Chart space not found.");
            C.Chart chart = chartSpace.GetFirstChild<C.Chart>() ??
                throw new InvalidOperationException("Chart not found.");
            bool cacheOnlySource = chartSpace.GetFirstChild<C.ExternalData>() == null;
            C.PlotArea plotArea = chart.GetFirstChild<C.PlotArea>() ??
                throw new InvalidOperationException("Chart plot area not found.");
            bool previousLegendUsesCategories = GetSharedNativeChartLayers(plotArea).FirstOrDefault() is
                C.PieChart or C.DoughnutChart;

            ISet<uint>? preservedSeriesIndexes = null;
            if (defaultKind == OfficeChartKind.Scatter && IsOnlySharedScatterPlot(plotArea)) {
                preservedSeriesIndexes = new HashSet<uint>(EnumerateSharedSeriesElements(plotArea)
                    .Select(series => series.GetFirstChild<C.Index>()?.Val?.Value ?? uint.MaxValue));
                UpdateScatterData(chartPart, NormalizeScatterData(data));
            } else {
                preservedSeriesIndexes = new HashSet<uint>();
                var replacement = new C.PlotArea();
                replacement.Append(plotArea.GetFirstChild<C.Layout>()?.CloneNode(true) ?? new C.Layout());
                AppendSharedChartContent(replacement, data, defaultKind);
                PreserveSharedChartFormatting(
                    plotArea, replacement, preservedSeriesIndexes);
                foreach (OpenXmlElement child in plotArea.ChildElements) {
                    if (child is C.DataTable || child is C.ChartShapeProperties || child is C.ExtensionList) {
                        replacement.Append(child.CloneNode(true));
                    }
                }
                chart.ReplaceChild(replacement, plotArea);
            }

            if (cacheOnlySource) {
                foreach (C.ValueAxis axis in chart.GetFirstChild<C.PlotArea>()!.Elements<C.ValueAxis>()) {
                    C.NumberingFormat? format = axis.GetFirstChild<C.NumberingFormat>();
                    if (format?.SourceLinked?.Value == true &&
                        string.Equals(format.FormatCode?.Value, "General", StringComparison.OrdinalIgnoreCase))
                        format.SourceLinked = false;
                }
            }

            UpdateSharedLegend(chart, data, defaultKind, previousLegendUsesCategories);
            ApplySharedChartSeriesStyle(chartPart, data, defaultKind, preservedSeriesIndexes);
            chartSpace.Save();
        }

        internal static void ApplySharedChartSeriesStyle(ChartPart chartPart, OfficeChartData data,
            OfficeChartKind defaultKind,
            ISet<uint>? preservedSeriesIndexes = null) {
            C.PlotArea? plotArea = chartPart.ChartSpace?.GetFirstChild<C.Chart>()?.GetFirstChild<C.PlotArea>();
            if (plotArea == null) return;
            foreach (OpenXmlCompositeElement seriesElement in EnumerateSharedSeriesElements(plotArea)) {
                uint nativeIndex =
                    seriesElement.GetFirstChild<C.Index>()?.Val?.Value ??
                    uint.MaxValue;
                int index = (int)nativeIndex;
                if (index < 0 || index >= data.Series.Count) continue;
                OfficeChartSeries series = data.Series[index];
                OfficeChartKind kind = series.RenderKind ?? defaultKind;
                bool newSeries = preservedSeriesIndexes == null || !preservedSeriesIndexes.Contains(nativeIndex);
                OfficeColor? fallbackSeriesColor = newSeries &&
                    kind is not OfficeChartKind.Pie and not OfficeChartKind.Doughnut &&
                    !series.Color.HasValue
                        ? OfficeChartStyle.Default.GetSeriesColor(index)
                        : null;
                ApplySharedSeriesShapeStyle(seriesElement, series, kind,
                    fallbackSeriesColor ?? (!newSeries && !series.Color.HasValue &&
                        seriesElement.GetFirstChild<C.ChartShapeProperties>()?.GetFirstChild<A.Outline>()?
                            .GetFirstChild<A.NoFill>() != null
                        ? ReadDirectSeriesColor(seriesElement) : null));
                ApplySharedSeriesMarker(seriesElement, series, kind, fallbackSeriesColor);
                ApplySharedPointColors(seriesElement, series);
                OfficeOpenXmlChartPointStyles.ApplySeries(seriesElement, series);
            }
        }

        private static void AppendSharedChartContent(C.PlotArea plotArea, OfficeChartData data,
            OfficeChartKind defaultKind) {
            List<SharedSeriesDescriptor> descriptors = DescribeSharedSeries(data, defaultKind);
            if (descriptors.All(item => item.Kind == OfficeChartKind.Pie)) {
                plotArea.Append(CreatePieChart(data));
                return;
            }
            if (descriptors.All(item => item.Kind == OfficeChartKind.Doughnut)) {
                plotArea.Append(CreateDoughnutChart(data));
                return;
            }
            if (descriptors.All(item => item.Kind == OfficeChartKind.Scatter)) {
                C.ScatterChart scatter = CreateScatterChart(NormalizeScatterData(data), out uint xAxisId, out uint yAxisId);
                plotArea.Append(scatter);
                plotArea.Append(CreateValueAxis(xAxisId, yAxisId, C.AxisPositionValues.Bottom));
                plotArea.Append(CreateValueAxis(yAxisId, xAxisId, C.AxisPositionValues.Left));
                return;
            }
            if (descriptors.All(item => item.Kind == OfficeChartKind.Bubble)) {
                uint xAxisId = GetNextAxisId();
                uint yAxisId = GetNextAxisId();
                plotArea.Append(CreateSharedBubbleChart(data, xAxisId, yAxisId));
                plotArea.Append(CreateSharedValueAxis(xAxisId, yAxisId,
                    C.AxisPositionValues.Bottom, secondary: false,
                    showMajorGridlines: false,
                    materializeDefaultGridlineStyle: false));
                plotArea.Append(CreateSharedValueAxis(yAxisId, xAxisId,
                    C.AxisPositionValues.Left, secondary: false,
                    showMajorGridlines: true,
                    materializeDefaultGridlineStyle: true));
                return;
            }

            bool horizontal = descriptors.All(item => IsHorizontalBarKind(item.Kind));
            bool hasPrimary = descriptors.Any(item => item.AxisGroup == OfficeChartAxisGroup.Primary);
            bool hasSecondary = descriptors.Any(item => item.AxisGroup == OfficeChartAxisGroup.Secondary);
            uint primaryCategoryId = hasPrimary ? GetNextAxisId() : 0U;
            uint primaryValueId = hasPrimary ? GetNextAxisId() : 0U;
            uint secondaryCategoryId = hasSecondary ? GetNextAxisId() : 0U;
            uint secondaryValueId = hasSecondary ? GetNextAxisId() : 0U;

            foreach (IGrouping<(OfficeChartKind Kind, OfficeChartAxisGroup AxisGroup), SharedSeriesDescriptor> group in
                     descriptors.GroupBy(item => (item.Kind, item.AxisGroup)).OrderBy(item => ChartLayer(item.Key.Kind))) {
                uint categoryId = group.Key.AxisGroup == OfficeChartAxisGroup.Primary
                    ? primaryCategoryId : secondaryCategoryId;
                uint valueId = group.Key.AxisGroup == OfficeChartAxisGroup.Primary
                    ? primaryValueId : secondaryValueId;
                List<SharedSeriesDescriptor> items = group.ToList();
                if (IsBarOrColumnKind(group.Key.Kind)) {
                    plotArea.Append(CreateSharedBarChart(group.Key.Kind, items, data.Categories, categoryId, valueId));
                } else if (IsLineKind(group.Key.Kind)) {
                    plotArea.Append(CreateSharedLineChart(group.Key.Kind, items, data.Categories, categoryId, valueId));
                } else if (IsAreaKind(group.Key.Kind)) {
                    plotArea.Append(CreateSharedAreaChart(group.Key.Kind, items, data.Categories, categoryId, valueId));
                } else if (group.Key.Kind == OfficeChartKind.Radar) {
                    plotArea.Append(CreateSharedRadarChart(items, data.Categories, categoryId, valueId));
                } else {
                    throw new NotSupportedException("Chart kind " + group.Key.Kind + " is not supported in this chart composition.");
                }
            }

            if (hasPrimary) {
                C.AxisPositionValues categoryPosition = horizontal
                    ? C.AxisPositionValues.Left : C.AxisPositionValues.Bottom;
                C.AxisPositionValues valuePosition = horizontal
                    ? C.AxisPositionValues.Bottom : C.AxisPositionValues.Left;
                plotArea.Append(CreateSharedCategoryAxis(primaryCategoryId, primaryValueId,
                    categoryPosition, secondary: false));
                plotArea.Append(CreateSharedValueAxis(primaryValueId, primaryCategoryId,
                    valuePosition, secondary: false,
                    showMajorGridlines: true,
                    materializeDefaultGridlineStyle: false));
            }
            if (hasSecondary) {
                plotArea.Append(CreateSharedCategoryAxis(secondaryCategoryId, secondaryValueId,
                    C.AxisPositionValues.Top, secondary: true));
                plotArea.Append(CreateSharedValueAxis(secondaryValueId, secondaryCategoryId,
                    C.AxisPositionValues.Right, secondary: true,
                    showMajorGridlines: false,
                    materializeDefaultGridlineStyle: false));
            }
        }

        private static C.BarChart CreateSharedBarChart(OfficeChartKind kind,
            IReadOnlyList<SharedSeriesDescriptor> descriptors, IReadOnlyList<string> categories,
            uint categoryAxisId, uint valueAxisId) {
            C.BarGroupingValues grouping = GetBarGrouping(kind);
            C.BarChart chart = new(
                new C.BarDirection { Val = IsHorizontalBarKind(kind) ? C.BarDirectionValues.Bar : C.BarDirectionValues.Column },
                new C.BarGrouping { Val = grouping },
                new C.VaryColors { Val = false });
            foreach (SharedSeriesDescriptor descriptor in descriptors) {
                chart.Append(CreateBarChartSeries(descriptor.Index,
                    descriptor.Series, categories));
            }
            chart.Append(CreateDefaultDataLabels());
            chart.Append(new C.GapWidth { Val = (UInt16Value)150U });
            chart.Append(new C.Overlap { Val = (SByteValue)(sbyte)(grouping == C.BarGroupingValues.Clustered ? 0 : 100) });
            chart.Append(new C.AxisId { Val = categoryAxisId });
            chart.Append(new C.AxisId { Val = valueAxisId });
            return chart;
        }

        private static C.LineChart CreateSharedLineChart(OfficeChartKind kind,
            IReadOnlyList<SharedSeriesDescriptor> descriptors, IReadOnlyList<string> categories,
            uint categoryAxisId, uint valueAxisId) {
            C.LineChart chart = new(new C.Grouping { Val = GetLineGrouping(kind) },
                new C.VaryColors { Val = false });
            foreach (SharedSeriesDescriptor descriptor in descriptors) {
                chart.Append(CreateLineChartSeries(descriptor.Index,
                    descriptor.Series, categories));
            }
            chart.Append(CreateDefaultDataLabels());
            chart.Append(new C.AxisId { Val = categoryAxisId });
            chart.Append(new C.AxisId { Val = valueAxisId });
            return chart;
        }

        private static C.AreaChart CreateSharedAreaChart(OfficeChartKind kind,
            IReadOnlyList<SharedSeriesDescriptor> descriptors, IReadOnlyList<string> categories,
            uint categoryAxisId, uint valueAxisId) {
            C.AreaChart chart = new(new C.Grouping { Val = GetAreaGrouping(kind) },
                new C.VaryColors { Val = false });
            foreach (SharedSeriesDescriptor descriptor in descriptors) {
                chart.Append(CreateSharedAreaSeries(descriptor, categories));
            }
            chart.Append(CreateDefaultDataLabels());
            chart.Append(new C.AxisId { Val = categoryAxisId });
            chart.Append(new C.AxisId { Val = valueAxisId });
            return chart;
        }

        private static C.RadarChart CreateSharedRadarChart(IReadOnlyList<SharedSeriesDescriptor> descriptors,
            IReadOnlyList<string> categories, uint categoryAxisId, uint valueAxisId) {
            C.RadarChart chart = new(new C.RadarStyle { Val = C.RadarStyleValues.Marker },
                new C.VaryColors { Val = false });
            foreach (SharedSeriesDescriptor descriptor in descriptors) {
                chart.Append(CreateSharedRadarSeries(descriptor, categories));
            }
            chart.Append(CreateDefaultDataLabels());
            chart.Append(new C.AxisId { Val = categoryAxisId });
            chart.Append(new C.AxisId { Val = valueAxisId });
            return chart;
        }

        private static C.AreaChartSeries CreateSharedAreaSeries(SharedSeriesDescriptor descriptor,
            IReadOnlyList<string> categories) {
            string column = ColumnLetter(descriptor.Index + 2);
            int lastRow = categories.Count + 1;
            return new C.AreaChartSeries(
                new C.Index { Val = (uint)descriptor.Index },
                new C.Order { Val = (uint)descriptor.Index },
                new C.SeriesText(CreateStringReference("Sheet1!$" + column + "$1",
                    new[] { descriptor.Series.Name })),
                new C.CategoryAxisData(CreateStringReference("Sheet1!$A$2:$A$" + lastRow, categories)),
                new C.Values(CreateNumberReference("Sheet1!$" + column + "$2:$" + column + "$" + lastRow,
                    descriptor.Series.Values)));
        }

        private static C.RadarChartSeries CreateSharedRadarSeries(SharedSeriesDescriptor descriptor,
            IReadOnlyList<string> categories) {
            string column = ColumnLetter(descriptor.Index + 2);
            int lastRow = categories.Count + 1;
            return new C.RadarChartSeries(
                new C.Index { Val = (uint)descriptor.Index },
                new C.Order { Val = (uint)descriptor.Index },
                new C.SeriesText(CreateStringReference("Sheet1!$" + column + "$1",
                    new[] { descriptor.Series.Name })),
                new C.CategoryAxisData(CreateStringReference("Sheet1!$A$2:$A$" + lastRow, categories)),
                new C.Values(CreateNumberReference("Sheet1!$" + column + "$2:$" + column + "$" + lastRow,
                    descriptor.Series.Values)));
        }

        private static C.CategoryAxis CreateSharedCategoryAxis(uint axisId, uint crossingAxisId,
            C.AxisPositionValues position, bool secondary) => new(
            new C.AxisId { Val = axisId },
            new C.Scaling(new C.Orientation { Val = C.OrientationValues.MinMax }),
            new C.Delete { Val = secondary },
            new C.AxisPosition { Val = position },
            new C.NumberingFormat { FormatCode = "General", SourceLinked = false },
            new C.MajorTickMark { Val = C.TickMarkValues.None },
            new C.MinorTickMark { Val = C.TickMarkValues.None },
            new C.TickLabelPosition { Val = secondary ? C.TickLabelPositionValues.None : C.TickLabelPositionValues.NextTo },
            new C.CrossingAxis { Val = crossingAxisId },
            new C.Crosses { Val = secondary ? C.CrossesValues.Maximum : C.CrossesValues.AutoZero },
            new C.AutoLabeled { Val = true },
            new C.LabelAlignment { Val = C.LabelAlignmentValues.Center },
            new C.LabelOffset { Val = (UInt16Value)100U },
            new C.NoMultiLevelLabels { Val = false });

        private static C.ValueAxis CreateSharedValueAxis(uint axisId, uint crossingAxisId,
            C.AxisPositionValues position, bool secondary,
            bool showMajorGridlines,
            bool materializeDefaultGridlineStyle) {
            C.ValueAxis axis = new(
                new C.AxisId { Val = axisId },
                new C.Scaling(new C.Orientation { Val = C.OrientationValues.MinMax }),
                new C.Delete { Val = false },
                new C.AxisPosition { Val = position });
            if (showMajorGridlines) {
                if (materializeDefaultGridlineStyle) {
                    OfficeChartStyle style = OfficeChartStyle.Default;
                    var outline = new A.Outline {
                        Width = checked((int)FromPoints(
                            style.GridLineWidth ?? 0.5D))
                    };
                    outline.Append(new A.SolidFill(
                        CreateSharedRgbColor(style.GridLineColor)));
                    axis.Append(new C.MajorGridlines(
                        new C.ChartShapeProperties(outline)));
                } else {
                    axis.Append(new C.MajorGridlines());
                }
            }
            axis.Append(new C.NumberingFormat { FormatCode = "General", SourceLinked = false });
            axis.Append(new C.MajorTickMark { Val = C.TickMarkValues.None });
            axis.Append(new C.MinorTickMark { Val = C.TickMarkValues.None });
            axis.Append(new C.TickLabelPosition { Val = C.TickLabelPositionValues.NextTo });
            axis.Append(new C.CrossingAxis { Val = crossingAxisId });
            axis.Append(new C.Crosses { Val = secondary ? C.CrossesValues.Maximum : C.CrossesValues.AutoZero });
            axis.Append(new C.CrossBetween { Val = C.CrossBetweenValues.Between });
            return axis;
        }

        private static IEnumerable<uint> GetHiddenSharedLegendIndexes(OfficeChartData data, OfficeChartKind kind) {
            kind = data.Series.FirstOrDefault()?.RenderKind ?? kind;
            if (kind == OfficeChartKind.Pie || kind == OfficeChartKind.Doughnut) {
                if (data.Series.Any(series => !series.ShowInLegend))
                    for (int point = 0; point < data.Categories.Count; point++) yield return (uint)point;
            } else {
                for (int index = 0; index < data.Series.Count; index++)
                    if (!data.Series[index].ShowInLegend) yield return (uint)index;
            }
        }

        internal static IEnumerable<OpenXmlElement> GetSharedNativeChartLayers(C.PlotArea plotArea) =>
            plotArea.ChildElements.Where(element => element.LocalName.EndsWith("Chart", StringComparison.OrdinalIgnoreCase));

        private static bool IsOnlySharedScatterPlot(C.PlotArea plotArea) {
            if (!plotArea.Elements<C.ScatterChart>().Any() ||
                !GetSharedNativeChartLayers(plotArea).All(element => element is C.ScatterChart)) return false;
            var axisCounts = plotArea.Elements<C.ValueAxis>().Where(axis => axis.AxisId?.Val != null)
                .GroupBy(axis => axis.AxisId!.Val!.Value).ToDictionary(group => group.Key, group => group.Count());
            uint[]? expected = null;
            foreach (C.ScatterChart layer in plotArea.Elements<C.ScatterChart>()) {
                var references = layer.Elements<C.AxisId>().Take(3).ToList();
                if (references.Count != 2 || references.Any(axis => axis.Val == null))
                    throw new NotSupportedException("In-place scatter updates require two referenced value axes per layer.");
                uint[] pair = references.Select(axis => axis.Val!.Value).ToArray();
                if (pair[0] == pair[1] || pair.Any(id => !axisCounts.TryGetValue(id, out int count) || count != 1) ||
                    expected != null && !expected.SequenceEqual(pair))
                    throw new NotSupportedException("In-place scatter updates require all layers to share the same ordered value-axis pair.");
                expected ??= pair;
            }
            return true;
        }

        private static C.Legend CreateSharedLegend(OfficeChartData data, OfficeChartKind kind) {
            C.Legend legend = new(new C.LegendPosition { Val = C.LegendPositionValues.Bottom });
            foreach (uint index in GetHiddenSharedLegendIndexes(data, kind)) {
                legend.Append(new C.LegendEntry(new C.Index { Val = index }, new C.Delete { Val = true }));
            }
            legend.Append(new C.Layout());
            legend.Append(new C.Overlay { Val = false });
            return legend;
        }

        private static void UpdateSharedLegend(C.Chart chart, OfficeChartData data, OfficeChartKind kind,
            bool previousLegendUsesCategories) {
            C.Legend? current = chart.GetFirstChild<C.Legend>();
            if (current == null) {
                return;
            }

            var replacement = (C.Legend)current.CloneNode(true);
            kind = data.Series.FirstOrDefault()?.RenderKind ?? kind;
            bool legendUsesCategories = kind is OfficeChartKind.Pie or OfficeChartKind.Doughnut;
            int entryCount = legendUsesCategories ? data.Categories.Count : data.Series.Count;
            var retained = (previousLegendUsesCategories == legendUsesCategories
                ? replacement.Elements<C.LegendEntry>() : Enumerable.Empty<C.LegendEntry>())
                .Where(entry => entry.GetFirstChild<C.Index>()?.Val?.Value < entryCount)
                .Where(entry => entry.ChildElements.Any(child => child is not C.Index && child is not C.Delete))
                .GroupBy(entry => entry.GetFirstChild<C.Index>()!.Val!.Value)
                .ToDictionary(group => group.Key, group => (C.LegendEntry)group.First().CloneNode(true));
            var hidden = new HashSet<uint>(GetHiddenSharedLegendIndexes(data, kind));
            replacement.RemoveAllChildren<C.LegendEntry>();
            OpenXmlElement? insertBefore = replacement.ChildElements.FirstOrDefault(child =>
                child is not C.LegendPosition && child is not C.LegendEntry);
            foreach (uint index in retained.Keys.Concat(hidden).Distinct().OrderBy(value => value)) {
                C.LegendEntry entry = retained.TryGetValue(index, out C.LegendEntry? existing)
                    ? existing : new C.LegendEntry(new C.Index { Val = index });
                entry.GetFirstChild<C.Delete>()?.Remove();
                if (hidden.Contains(index)) entry.AddChild(new C.Delete { Val = true }, true);
                if (insertBefore == null) replacement.Append(entry);
                else replacement.InsertBefore(entry, insertBefore);
            }
            chart.ReplaceChild(replacement, current);
        }

        private static IEnumerable<OpenXmlCompositeElement> EnumerateSharedSeriesElements(C.PlotArea plotArea) {
            foreach (OpenXmlElement chart in plotArea.ChildElements) {
                if (chart is C.BarChart bar) foreach (C.BarChartSeries series in bar.Elements<C.BarChartSeries>()) yield return series;
                else if (chart is C.LineChart line) foreach (C.LineChartSeries series in line.Elements<C.LineChartSeries>()) yield return series;
                else if (chart is C.AreaChart area) foreach (C.AreaChartSeries series in area.Elements<C.AreaChartSeries>()) yield return series;
                else if (chart is C.RadarChart radar) foreach (C.RadarChartSeries series in radar.Elements<C.RadarChartSeries>()) yield return series;
                else if (chart is C.ScatterChart scatter) foreach (C.ScatterChartSeries series in scatter.Elements<C.ScatterChartSeries>()) yield return series;
                else if (chart is C.BubbleChart bubble) foreach (C.BubbleChartSeries series in bubble.Elements<C.BubbleChartSeries>()) yield return series;
                else if (chart is C.PieChart pie) foreach (C.PieChartSeries series in pie.Elements<C.PieChartSeries>()) yield return series;
                else if (chart is C.DoughnutChart doughnut) foreach (C.PieChartSeries series in doughnut.Elements<C.PieChartSeries>()) yield return series;
            }
        }

        private static List<SharedSeriesDescriptor> DescribeSharedSeries(OfficeChartData data,
            OfficeChartKind defaultKind) => data.Series.Select((series, index) =>
                new SharedSeriesDescriptor(index, series, series.RenderKind ?? defaultKind)).ToList();

        private static IReadOnlyList<double> ParseScatterCategories(IReadOnlyList<string> categories) {
            var values = new List<double>(categories.Count);
            foreach (string category in categories) {
                if (!double.TryParse(category, NumberStyles.Float, CultureInfo.InvariantCulture, out double value) ||
                    double.IsNaN(value) || double.IsInfinity(value)) {
                    throw new ArgumentException(
                        "Scatter chart categories must be finite invariant numeric values when a series has no XValues.",
                        nameof(categories));
                }
                values.Add(value);
            }
            return values;
        }

        private static bool IsBarOrColumnKind(OfficeChartKind kind) =>
            kind == OfficeChartKind.ColumnClustered || kind == OfficeChartKind.ColumnStacked ||
            kind == OfficeChartKind.ColumnStacked100 || kind == OfficeChartKind.BarClustered ||
            kind == OfficeChartKind.BarStacked || kind == OfficeChartKind.BarStacked100;

        private static bool IsHorizontalBarKind(OfficeChartKind kind) =>
            kind == OfficeChartKind.BarClustered || kind == OfficeChartKind.BarStacked ||
            kind == OfficeChartKind.BarStacked100;

        private static bool IsLineKind(OfficeChartKind kind) =>
            kind == OfficeChartKind.Line || kind == OfficeChartKind.LineStacked ||
            kind == OfficeChartKind.LineStacked100;

        private static bool IsAreaKind(OfficeChartKind kind) =>
            kind == OfficeChartKind.Area || kind == OfficeChartKind.AreaStacked ||
            kind == OfficeChartKind.AreaStacked100;

        private static bool IsFilledSharedKind(OfficeChartKind kind) =>
            IsBarOrColumnKind(kind) || IsAreaKind(kind) || kind == OfficeChartKind.Pie ||
            kind == OfficeChartKind.Doughnut || kind == OfficeChartKind.Bubble;

        private static bool IsMarkerKind(OfficeChartKind kind) => IsLineKind(kind) ||
            kind == OfficeChartKind.Scatter || kind == OfficeChartKind.Radar;

        private static C.BarGroupingValues GetBarGrouping(OfficeChartKind kind) {
            if (kind == OfficeChartKind.ColumnStacked100 || kind == OfficeChartKind.BarStacked100)
                return C.BarGroupingValues.PercentStacked;
            if (kind == OfficeChartKind.ColumnStacked || kind == OfficeChartKind.BarStacked)
                return C.BarGroupingValues.Stacked;
            return C.BarGroupingValues.Clustered;
        }

        private static C.GroupingValues GetLineGrouping(OfficeChartKind kind) {
            if (kind == OfficeChartKind.LineStacked100) return C.GroupingValues.PercentStacked;
            if (kind == OfficeChartKind.LineStacked) return C.GroupingValues.Stacked;
            return C.GroupingValues.Standard;
        }

        private static C.GroupingValues GetAreaGrouping(OfficeChartKind kind) {
            if (kind == OfficeChartKind.AreaStacked100) return C.GroupingValues.PercentStacked;
            if (kind == OfficeChartKind.AreaStacked) return C.GroupingValues.Stacked;
            return C.GroupingValues.Standard;
        }

        private static int ChartLayer(OfficeChartKind kind) {
            if (IsAreaKind(kind)) return 0;
            if (IsBarOrColumnKind(kind)) return 1;
            return 2;
        }

        private static C.MarkerStyleValues MapMarker(OfficeChartMarkerShape? marker) {
            if (!marker.HasValue) return C.MarkerStyleValues.Circle;
            switch (marker.Value) {
                case OfficeChartMarkerShape.Square: return C.MarkerStyleValues.Square;
                case OfficeChartMarkerShape.Diamond: return C.MarkerStyleValues.Diamond;
                case OfficeChartMarkerShape.Triangle: return C.MarkerStyleValues.Triangle;
                case OfficeChartMarkerShape.Dash: return C.MarkerStyleValues.Dash;
                case OfficeChartMarkerShape.Dot: return C.MarkerStyleValues.Dot;
                case OfficeChartMarkerShape.Plus: return C.MarkerStyleValues.Plus;
                case OfficeChartMarkerShape.X: return C.MarkerStyleValues.X;
                case OfficeChartMarkerShape.Star: return C.MarkerStyleValues.Star;
                default: return C.MarkerStyleValues.Circle;
            }
        }

        private static A.PresetLineDashValues MapDash(OfficeStrokeDashStyle dash) {
            switch (dash) {
                case OfficeStrokeDashStyle.Dash: return A.PresetLineDashValues.Dash;
                case OfficeStrokeDashStyle.Dot: return A.PresetLineDashValues.Dot;
                case OfficeStrokeDashStyle.DashDot: return A.PresetLineDashValues.DashDot;
                case OfficeStrokeDashStyle.DashDotDot: return A.PresetLineDashValues.LargeDashDotDot;
                default: return A.PresetLineDashValues.Solid;
            }
        }
    }
}
