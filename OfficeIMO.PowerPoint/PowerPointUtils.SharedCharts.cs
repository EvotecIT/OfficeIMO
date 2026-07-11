using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Drawing;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.PowerPoint {
    internal static partial class PowerPointUtils {
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

        internal static void ValidateSharedChartData(OfficeChartData data, OfficeChartKind defaultKind) {
            if (data == null) throw new ArgumentNullException(nameof(data));
            for (int index = 0; index < data.Series.Count; index++) {
                OfficeChartSeries series = data.Series[index];
                if (series.Values.Count == 0) {
                    throw new ArgumentException("Chart series cannot be empty.", nameof(data));
                }
                if (defaultKind == OfficeChartKind.Scatter) {
                    if (series.XValues != null && series.XValues.Count != series.Values.Count) {
                        throw new ArgumentException("Scatter X and Y value counts must match.", nameof(data));
                    }
                } else if (series.Values.Count != data.Categories.Count) {
                    throw new ArgumentException("Every chart series must match the category count.", nameof(data));
                }
            }

            List<SharedSeriesDescriptor> descriptors = DescribeSharedSeries(data, defaultKind);
            bool hasSecondary = descriptors.Any(item => item.AxisGroup == OfficeChartAxisGroup.Secondary);
            if (hasSecondary && descriptors.All(item => item.AxisGroup == OfficeChartAxisGroup.Secondary)) {
                throw new NotSupportedException("A secondary-axis chart requires at least one primary-axis series.");
            }

            if (descriptors.Any(item => item.Kind == OfficeChartKind.Scatter)) {
                if (defaultKind != OfficeChartKind.Scatter ||
                    descriptors.Any(item => item.Kind != OfficeChartKind.Scatter) || hasSecondary) {
                    throw new NotSupportedException("Scatter series cannot be combined with categorical or secondary-axis series.");
                }
                foreach (OfficeChartSeries series in data.Series) {
                    if (series.XValues == null) ParseScatterCategories(data.Categories);
                }
                return;
            }

            bool hasHorizontalBar = descriptors.Any(item => IsHorizontalBarKind(item.Kind));
            if (hasHorizontalBar && (descriptors.Any(item => !IsHorizontalBarKind(item.Kind)) || hasSecondary)) {
                throw new NotSupportedException("Horizontal bar charts cannot be mixed with other families or secondary axes.");
            }

            bool hasStandalone = descriptors.Any(item => item.Kind == OfficeChartKind.Pie ||
                item.Kind == OfficeChartKind.Doughnut || item.Kind == OfficeChartKind.Radar);
            if (hasStandalone && (descriptors.Select(item => item.Kind).Distinct().Count() > 1 || hasSecondary)) {
                throw new NotSupportedException("Pie, doughnut, and radar charts cannot participate in combo or secondary-axis charts.");
            }
        }

        internal static PowerPointChartData ToPowerPointChartData(OfficeChartData data) =>
            new(data.Categories, data.Series.Select(series =>
                new PowerPointChartSeries(series.Name, series.Values)));

        internal static PowerPointScatterChartData ToPowerPointScatterChartData(OfficeChartData data) {
            IReadOnlyList<double>? sharedX = data.Series.Any(series => series.XValues == null)
                ? ParseScatterCategories(data.Categories)
                : null;
            return new PowerPointScatterChartData(data.Series.Select(series =>
                new PowerPointScatterChartSeries(series.Name, series.XValues ?? sharedX!, series.Values)));
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
            chart.Append(CreateSharedLegend(data));
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
            C.PlotArea plotArea = chart.GetFirstChild<C.PlotArea>() ??
                throw new InvalidOperationException("Chart plot area not found.");

            if (defaultKind == OfficeChartKind.Scatter) {
                UpdateChartData(chartPart, ToPowerPointScatterChartData(data));
            } else {
                var replacement = new C.PlotArea();
                replacement.Append(plotArea.GetFirstChild<C.Layout>()?.CloneNode(true) ?? new C.Layout());
                AppendSharedChartContent(replacement, data, defaultKind);
                foreach (OpenXmlElement child in plotArea.ChildElements) {
                    if (child is C.DataTable || child is C.ChartShapeProperties || child is C.ExtensionList) {
                        replacement.Append(child.CloneNode(true));
                    }
                }
                chart.ReplaceChild(replacement, plotArea);
            }

            UpdateSharedLegend(chart, data);
            ApplySharedChartSeriesStyle(chartPart, data, defaultKind);
            chartSpace.Save();
        }

        internal static void ApplySharedChartSeriesStyle(ChartPart chartPart, OfficeChartData data,
            OfficeChartKind defaultKind) {
            C.PlotArea? plotArea = chartPart.ChartSpace?.GetFirstChild<C.Chart>()?.GetFirstChild<C.PlotArea>();
            if (plotArea == null) return;
            foreach (OpenXmlCompositeElement seriesElement in EnumerateSharedSeriesElements(plotArea)) {
                int index = (int)(seriesElement.GetFirstChild<C.Index>()?.Val?.Value ?? uint.MaxValue);
                if (index < 0 || index >= data.Series.Count) continue;
                OfficeChartSeries series = data.Series[index];
                OfficeChartKind kind = series.RenderKind ?? defaultKind;
                ApplySharedSeriesShapeStyle(seriesElement, series, kind);
                ApplySharedSeriesMarker(seriesElement, series, kind);
                ApplySharedPointColors(seriesElement, series);
            }
        }

        private static void AppendSharedChartContent(C.PlotArea plotArea, OfficeChartData data,
            OfficeChartKind defaultKind) {
            List<SharedSeriesDescriptor> descriptors = DescribeSharedSeries(data, defaultKind);
            if (descriptors.All(item => item.Kind == OfficeChartKind.Pie)) {
                plotArea.Append(CreatePieChart(ToPowerPointChartData(data)));
                return;
            }
            if (descriptors.All(item => item.Kind == OfficeChartKind.Doughnut)) {
                plotArea.Append(CreateDoughnutChart(ToPowerPointChartData(data)));
                return;
            }

            bool horizontal = descriptors.All(item => IsHorizontalBarKind(item.Kind));
            bool hasPrimary = descriptors.Any(item => item.AxisGroup == OfficeChartAxisGroup.Primary);
            bool hasSecondary = descriptors.Any(item => item.AxisGroup == OfficeChartAxisGroup.Secondary);
            uint primaryCategoryId = hasPrimary ? PowerPointChartAxisIdGenerator.GetNextId() : 0U;
            uint primaryValueId = hasPrimary ? PowerPointChartAxisIdGenerator.GetNextId() : 0U;
            uint secondaryCategoryId = hasSecondary ? PowerPointChartAxisIdGenerator.GetNextId() : 0U;
            uint secondaryValueId = hasSecondary ? PowerPointChartAxisIdGenerator.GetNextId() : 0U;

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
                    valuePosition, secondary: false));
            }
            if (hasSecondary) {
                plotArea.Append(CreateSharedCategoryAxis(secondaryCategoryId, secondaryValueId,
                    C.AxisPositionValues.Top, secondary: true));
                plotArea.Append(CreateSharedValueAxis(secondaryValueId, secondaryCategoryId,
                    C.AxisPositionValues.Right, secondary: true));
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
                    new PowerPointChartSeries(descriptor.Series.Name, descriptor.Series.Values), categories));
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
                    new PowerPointChartSeries(descriptor.Series.Name, descriptor.Series.Values), categories));
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
            new C.NumberingFormat { FormatCode = "General", SourceLinked = true },
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
            C.AxisPositionValues position, bool secondary) {
            C.ValueAxis axis = new(
                new C.AxisId { Val = axisId },
                new C.Scaling(new C.Orientation { Val = C.OrientationValues.MinMax }),
                new C.Delete { Val = false },
                new C.AxisPosition { Val = position });
            if (!secondary) axis.Append(new C.MajorGridlines());
            axis.Append(new C.NumberingFormat { FormatCode = "General", SourceLinked = true });
            axis.Append(new C.MajorTickMark { Val = C.TickMarkValues.None });
            axis.Append(new C.MinorTickMark { Val = C.TickMarkValues.None });
            axis.Append(new C.TickLabelPosition { Val = C.TickLabelPositionValues.NextTo });
            axis.Append(new C.CrossingAxis { Val = crossingAxisId });
            axis.Append(new C.Crosses { Val = secondary ? C.CrossesValues.Maximum : C.CrossesValues.AutoZero });
            axis.Append(new C.CrossBetween { Val = C.CrossBetweenValues.Between });
            return axis;
        }

        private static C.Legend CreateSharedLegend(OfficeChartData data) {
            C.Legend legend = new(new C.LegendPosition { Val = C.LegendPositionValues.Bottom });
            for (int index = 0; index < data.Series.Count; index++) {
                if (!data.Series[index].ShowInLegend) {
                    legend.Append(new C.LegendEntry(new C.Index { Val = (uint)index },
                        new C.Delete { Val = true }));
                }
            }
            legend.Append(new C.Layout());
            legend.Append(new C.Overlay { Val = false });
            return legend;
        }

        private static void UpdateSharedLegend(C.Chart chart, OfficeChartData data) {
            C.Legend? current = chart.GetFirstChild<C.Legend>();
            if (current == null) {
                C.PlotArea? plotArea = chart.GetFirstChild<C.PlotArea>();
                C.Legend created = CreateSharedLegend(data);
                if (plotArea != null) chart.InsertAfter(created, plotArea);
                else chart.Append(created);
                return;
            }

            var replacement = (C.Legend)current.CloneNode(true);
            replacement.RemoveAllChildren<C.LegendEntry>();
            OpenXmlElement? insertBefore = replacement.ChildElements.FirstOrDefault(child =>
                child is not C.LegendPosition && child is not C.LegendEntry);
            for (int index = 0; index < data.Series.Count; index++) {
                if (data.Series[index].ShowInLegend) continue;
                var entry = new C.LegendEntry(new C.Index { Val = (uint)index },
                    new C.Delete { Val = true });
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
                else if (chart is C.PieChart pie) foreach (C.PieChartSeries series in pie.Elements<C.PieChartSeries>()) yield return series;
                else if (chart is C.DoughnutChart doughnut) foreach (C.PieChartSeries series in doughnut.Elements<C.PieChartSeries>()) yield return series;
            }
        }

        private static void ApplySharedSeriesShapeStyle(OpenXmlCompositeElement seriesElement,
            OfficeChartSeries series, OfficeChartKind kind) {
            if (!series.Color.HasValue && series.StrokeWidth == null && series.StrokeDashStyle == null &&
                series.ConnectLine) return;
            C.ChartShapeProperties properties = seriesElement.GetFirstChild<C.ChartShapeProperties>() ??
                new C.ChartShapeProperties();
            if (series.Color.HasValue && IsFilledSharedKind(kind)) {
                properties.RemoveAllChildren<A.SolidFill>();
                properties.RemoveAllChildren<A.NoFill>();
                properties.PrependChild(new A.SolidFill(
                    new A.RgbColorModelHex { Val = series.Color.Value.ToRgbHex() }));
            }
            A.Outline outline = properties.GetFirstChild<A.Outline>() ?? new A.Outline();
            outline.RemoveAllChildren<A.SolidFill>();
            outline.RemoveAllChildren<A.NoFill>();
            if (!series.ConnectLine && !IsFilledSharedKind(kind)) {
                outline.Append(new A.NoFill());
            } else if (series.Color.HasValue) {
                outline.Append(new A.SolidFill(
                    new A.RgbColorModelHex { Val = series.Color.Value.ToRgbHex() }));
            }
            if (series.StrokeWidth.HasValue) {
                outline.Width = (int)Math.Min(int.MaxValue,
                    PowerPointUnits.FromPoints(series.StrokeWidth.Value));
            }
            if (series.StrokeDashStyle.HasValue) {
                outline.RemoveAllChildren<A.PresetDash>();
                outline.Append(new A.PresetDash { Val = MapDash(series.StrokeDashStyle.Value) });
            }
            if (outline.Parent == null) properties.Append(outline);
            if (properties.Parent == null) InsertSharedSeriesProperties(seriesElement, properties);
        }

        private static void ApplySharedSeriesMarker(OpenXmlCompositeElement seriesElement,
            OfficeChartSeries series, OfficeChartKind kind) {
            if (!IsMarkerKind(kind)) return;
            C.Marker marker = seriesElement.GetFirstChild<C.Marker>() ?? new C.Marker();
            marker.Symbol = new C.Symbol {
                Val = series.ShowMarkers ? MapMarker(series.MarkerShape) : C.MarkerStyleValues.None
            };
            if (series.MarkerSize.HasValue) marker.Size = new C.Size { Val = (byte)Math.Min(72, series.MarkerSize.Value) };
            if (series.Color.HasValue || series.MarkerOutlineColor.HasValue || series.MarkerOutlineWidth.HasValue) {
                C.ChartShapeProperties properties = marker.ChartShapeProperties ?? new C.ChartShapeProperties();
                if (series.Color.HasValue) {
                    properties.RemoveAllChildren<A.SolidFill>();
                    properties.PrependChild(new A.SolidFill(
                        new A.RgbColorModelHex { Val = series.Color.Value.ToRgbHex() }));
                }
                A.Outline outline = properties.GetFirstChild<A.Outline>() ?? new A.Outline();
                OfficeColor? markerColor = series.MarkerOutlineColor ?? series.Color;
                if (markerColor.HasValue) {
                    outline.RemoveAllChildren<A.SolidFill>();
                    outline.Append(new A.SolidFill(
                        new A.RgbColorModelHex { Val = markerColor.Value.ToRgbHex() }));
                }
                if (series.MarkerOutlineWidth.HasValue) {
                    outline.Width = (int)Math.Min(int.MaxValue,
                        PowerPointUnits.FromPoints(series.MarkerOutlineWidth.Value));
                }
                if (outline.Parent == null) properties.Append(outline);
                if (properties.Parent == null) marker.Append(properties);
            }
            if (marker.Parent == null) InsertSharedMarker(seriesElement, marker);
        }

        private static void ApplySharedPointColors(OpenXmlCompositeElement seriesElement, OfficeChartSeries series) {
            if (series.PointColors == null) return;
            seriesElement.RemoveAllChildren<C.DataPoint>();
            OpenXmlElement? insertBefore = seriesElement.GetFirstChild<C.DataLabels>() ??
                (OpenXmlElement?)seriesElement.GetFirstChild<C.CategoryAxisData>() ??
                (OpenXmlElement?)seriesElement.GetFirstChild<C.Values>() ??
                (OpenXmlElement?)seriesElement.GetFirstChild<C.XValues>() ?? seriesElement.GetFirstChild<C.YValues>();
            for (int index = 0; index < series.PointColors.Count; index++) {
                OfficeColor? color = series.PointColors[index];
                if (!color.HasValue) continue;
                C.DataPoint point = new(new C.Index { Val = (uint)index },
                    new C.ChartShapeProperties(new A.SolidFill(
                        new A.RgbColorModelHex { Val = color.Value.ToRgbHex() })));
                if (insertBefore != null) seriesElement.InsertBefore(point, insertBefore);
                else seriesElement.Append(point);
            }
        }

        private static void InsertSharedSeriesProperties(OpenXmlCompositeElement series, C.ChartShapeProperties properties) {
            OpenXmlElement? insertBefore = series.GetFirstChild<C.InvertIfNegative>() ??
                (OpenXmlElement?)series.GetFirstChild<C.Marker>() ?? series.GetFirstChild<C.DataPoint>() ??
                series.GetFirstChild<C.DataLabels>() ?? series.GetFirstChild<C.CategoryAxisData>() ??
                (OpenXmlElement?)series.GetFirstChild<C.Values>() ??
                (OpenXmlElement?)series.GetFirstChild<C.XValues>() ?? series.GetFirstChild<C.YValues>();
            if (insertBefore != null) series.InsertBefore(properties, insertBefore);
            else series.Append(properties);
        }

        private static void InsertSharedMarker(OpenXmlCompositeElement series, C.Marker marker) {
            OpenXmlElement? insertBefore = series.GetFirstChild<C.DataPoint>() ??
                (OpenXmlElement?)series.GetFirstChild<C.DataLabels>() ?? series.GetFirstChild<C.CategoryAxisData>() ??
                (OpenXmlElement?)series.GetFirstChild<C.Values>() ??
                (OpenXmlElement?)series.GetFirstChild<C.XValues>() ?? series.GetFirstChild<C.YValues>();
            if (insertBefore != null) series.InsertBefore(marker, insertBefore);
            else series.Append(marker);
        }

        private static List<SharedSeriesDescriptor> DescribeSharedSeries(OfficeChartData data,
            OfficeChartKind defaultKind) => data.Series.Select((series, index) =>
                new SharedSeriesDescriptor(index, series, series.RenderKind ?? defaultKind)).ToList();

        private static IReadOnlyList<double> ParseScatterCategories(IReadOnlyList<string> categories) {
            var values = new List<double>(categories.Count);
            foreach (string category in categories) {
                if (!double.TryParse(category, NumberStyles.Float, CultureInfo.InvariantCulture, out double value)) {
                    throw new ArgumentException(
                        "Scatter chart categories must be invariant numeric values when a series has no XValues.",
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
            kind == OfficeChartKind.Doughnut;

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
