using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.Drawing;
using static OfficeIMO.OpenXml.Internal.OfficeChartXml;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.OpenXml.Internal {
    internal static partial class OfficeOpenXmlChartWriter {
        private const int SpreadsheetMaximumColumns = 16_384;
        private const int SpreadsheetMaximumRows = 1_048_576;
        private const int BubbleWorkbookColumnsPerSeries = 3;
        internal const int MaximumSharedChartPoints = 100_000;

        private sealed class SharedSeriesDescriptor {
            internal SharedSeriesDescriptor(int index, OfficeSeries series, OfficeChartType kind) {
                Index = index;
                Series = series;
                Kind = kind;
            }

            internal int Index { get; }
            internal OfficeSeries Series { get; }
            internal OfficeChartType Kind { get; }
            internal OfficeChartAxisGroup AxisGroup => Series.AxisGroup;
        }

        /// <summary>
        /// Writes a new chart into the part's chart space root, which the host provides (an SDK host's forwarder makes a
        /// fresh one). Its child elements are replaced.
        /// </summary>
        internal static void PopulateSharedChart(IOfficeChartPart chartPart, string embeddedRelId, OfficeChartData data,
            OfficeChartType defaultKind, OfficeTheme? theme = null) {
            ArgumentNullException.ThrowIfNull(chartPart);
            ValidateSharedChartData(data, defaultKind);
            PopulateSharedChartCore(chartPart, embeddedRelId, data, defaultKind, theme);
            chartPart.MarkEdited();
        }

        private static void PopulateSharedChartCore(IOfficeChartPart chartPart, string embeddedRelId, OfficeChartData data,
            OfficeChartType defaultKind, OfficeTheme? theme) {
            OfficeXmlNode chartSpace = chartPart.ChartSpace;
            if (chartSpace.IsNone) throw new InvalidOperationException("Chart space not found.");
            chartSpace.RemoveAllChildElements();
            if (!chartSpace.IsSdk) {
                DeclareIfMissing(chartSpace, "c", ChartNamespace);
                DeclareIfMissing(chartSpace, "a", DrawingNamespace);
                DeclareIfMissing(chartSpace, "r", RelationshipNamespace);
            }
            chartSpace.AppendVal(NsC, "date1904", Text(false));
            chartSpace.AppendVal(NsC, "lang", "en-US");
            chartSpace.AppendVal(NsC, "roundedCorners", Text(false));

            // Written without a title: autoTitleDeleted stops Office drawing an automatic one. Setting a title clears it.
            OfficeXmlNode chart = chartSpace.AppendNew(NsC, "chart");
            chart.AppendVal(NsC, "autoTitleDeleted", Text(true));
            OfficeXmlNode plotArea = chart.AppendNew(NsC, "plotArea");
            plotArea.AppendNew(NsC, "layout");
            AppendSharedChartContent(plotArea, data, defaultKind);
            chart.Append(CreateSharedLegend(chart, data, defaultKind));
            chart.AppendVal(NsC, "plotVisOnly", Text(true));
            chart.AppendVal(NsC, "dispBlanksAs", "gap");
            chart.AppendVal(NsC, "showDLblsOverMax", Text(false));
            if (!string.IsNullOrWhiteSpace(embeddedRelId)) {
                OfficeXmlNode externalData = chartSpace.AppendNew(NsC, "externalData");
                externalData.SetAttribute(NsR, "id", embeddedRelId);
                externalData.AppendVal(NsC, "autoUpdate", Text(false));
            }
            ApplySharedChartSeriesStyle(chartPart, data, defaultKind, theme: theme);
        }

        private static void DeclareIfMissing(OfficeXmlNode element, string prefix, string namespaceUri) {
            if (element.Attribute(OfficeXmlNamespaces.Xmlns, prefix) == null) element.DeclareNamespace(prefix, namespaceUri);
        }

        internal static void UpdateSharedChartData(IOfficeChartPart chartPart, OfficeChartData data,
            OfficeChartType defaultKind, OfficeTheme? theme = null) {
            ArgumentNullException.ThrowIfNull(chartPart);
            ValidateSharedChartData(data, defaultKind);

            OfficeXmlNode chartSpace = chartPart.ChartSpace;
            if (chartSpace.IsNone) throw new InvalidOperationException("Chart space not found.");
            OfficeXmlNode chart = chartSpace.Child(NsC, "chart");
            if (chart.IsNone) throw new InvalidOperationException("Chart not found.");
            bool cacheOnlySource = chartSpace.Child(NsC, "externalData").IsNone;
            OfficeXmlNode plotArea = chart.Child(NsC, "plotArea");
            if (plotArea.IsNone) throw new InvalidOperationException("Chart plot area not found.");
            OfficeXmlNode firstLayer = GetSharedNativeChartLayers(plotArea).FirstOrDefault();
            bool previousLegendUsesCategories = firstLayer.Is(NsC, "pieChart") || firstLayer.Is(NsC, "doughnutChart");

            ISet<uint>? preservedSeriesIndexes = null;
            if (defaultKind == OfficeChartType.Scatter && IsOnlySharedScatterPlot(plotArea)) {
                preservedSeriesIndexes = new HashSet<uint>(EnumerateSharedSeriesElements(plotArea)
                    .Select(series => series.Child(NsC, "idx").UIntVal() ?? uint.MaxValue));
                UpdateScatterData(chartPart, NormalizeScatterData(data));
            } else {
                preservedSeriesIndexes = new HashSet<uint>();
                OfficeXmlNode replacement = chart.CreateElement(NsC, "plotArea");
                OfficeXmlNode layout = plotArea.Child(NsC, "layout");
                replacement.Append(layout.IsNone ? replacement.CreateElement(NsC, "layout") : layout.Clone());
                AppendSharedChartContent(replacement, data, defaultKind);
                PreserveSharedChartFormatting(
                    plotArea, replacement, preservedSeriesIndexes);
                foreach (OfficeXmlNode child in plotArea.Children()) {
                    if (child.Is(NsC, "dTable") || child.Is(NsC, "spPr") || child.Is(NsC, "extLst")) {
                        replacement.Append(child.Clone());
                    }
                }
                plotArea.ReplaceWith(replacement);
            }

            if (cacheOnlySource) {
                foreach (OfficeXmlNode axis in chart.Child(NsC, "plotArea").Children(NsC, "valAx")) {
                    OfficeXmlNode format = axis.Child(NsC, "numFmt");
                    if (!format.IsNone && format.BoolAttr("sourceLinked") == true &&
                        string.Equals(format.Attribute("formatCode"), "General", StringComparison.OrdinalIgnoreCase))
                        format.SetAttribute("sourceLinked", Text(false));
                }
            }

            UpdateSharedLegend(chart, data, defaultKind, previousLegendUsesCategories);
            ApplySharedChartSeriesStyle(chartPart, data, defaultKind, preservedSeriesIndexes, theme);
            chartPart.MarkEdited();
        }

        internal static void ApplySharedChartSeriesStyle(IOfficeChartPart chartPart, OfficeChartData data,
            OfficeChartType defaultKind,
            ISet<uint>? preservedSeriesIndexes = null, OfficeTheme? theme = null) {
            theme ??= OfficeTheme.Default;
            OfficeXmlNode plotArea = chartPart.ChartSpace.Child(NsC, "chart").Child(NsC, "plotArea");
            if (plotArea.IsNone) return;
            foreach (OfficeXmlNode seriesElement in EnumerateSharedSeriesElements(plotArea)) {
                uint nativeIndex =
                    seriesElement.Child(NsC, "idx").UIntVal() ??
                    uint.MaxValue;
                int index = (int)nativeIndex;
                if (index < 0 || index >= data.Series.Count) continue;
                OfficeSeries series = data.Series[index];
                OfficeChartType kind = series.RenderKind ?? defaultKind;
                bool newSeries = preservedSeriesIndexes == null || !preservedSeriesIndexes.Contains(nativeIndex);
                OfficeColor? fallbackSeriesColor = newSeries &&
                    kind is not OfficeChartType.Pie and not OfficeChartType.Doughnut &&
                    !series.Color.HasValue
                        ? theme.ChartStyle.SeriesPalette[index % theme.ChartStyle.SeriesPalette.Count].Resolve(theme)
                        : null;
                ApplySharedSeriesShapeStyle(seriesElement, series, kind,
                    fallbackSeriesColor ?? (!newSeries && !series.Color.HasValue &&
                        !seriesElement.Child(NsC, "spPr").Child(NsA, "ln").Child(NsA, "noFill").IsNone
                        ? ReadDirectSeriesColor(seriesElement) : null), theme);
                if (fallbackSeriesColor.HasValue || series.Color.HasValue)
                    SetAutomaticSeriesColor(seriesElement, fallbackSeriesColor.HasValue);
                ApplySharedSeriesMarker(seriesElement, series, kind, fallbackSeriesColor, theme,
                    preservePresentation: !newSeries && !series.HasExplicitMarkerVisibility);
                ApplySharedPointColors(seriesElement, series, theme);
                OfficeOpenXmlChartPointStyles.ApplySeries(seriesElement, series, theme);
                if (series.SeriesStyle?.FillColor.HasValue == true || series.SeriesStyle?.OutlineColor.HasValue == true)
                    SetAutomaticSeriesColor(seriesElement, false);
                OfficeOpenXmlChartExplosions.ApplySeries(seriesElement, series);
            }
        }

        private static void AppendSharedChartContent(OfficeXmlNode plotArea, OfficeChartData data,
            OfficeChartType defaultKind) {
            List<SharedSeriesDescriptor> descriptors = DescribeSharedSeries(data, defaultKind);
            if (descriptors.All(item => item.Kind == OfficeChartType.Pie)) {
                plotArea.Append(CreatePieChart(plotArea, data));
                return;
            }
            if (descriptors.All(item => item.Kind == OfficeChartType.Doughnut)) {
                plotArea.Append(CreateDoughnutChart(plotArea, data));
                return;
            }
            if (descriptors.All(item => item.Kind == OfficeChartType.Scatter)) {
                OfficeXmlNode scatter = CreateScatterChart(plotArea, NormalizeScatterData(data), out uint xAxisId, out uint yAxisId);
                plotArea.Append(scatter);
                plotArea.Append(CreateValueAxis(plotArea, xAxisId, yAxisId, "b"));
                plotArea.Append(CreateValueAxis(plotArea, yAxisId, xAxisId, "l"));
                return;
            }
            if (descriptors.All(item => item.Kind == OfficeChartType.Bubble)) {
                uint xAxisId = GetNextAxisId();
                uint yAxisId = GetNextAxisId();
                plotArea.Append(CreateSharedBubbleChart(plotArea, data, xAxisId, yAxisId));
                plotArea.Append(CreateSharedValueAxis(plotArea, xAxisId, yAxisId,
                    "b", secondary: false,
                    showMajorGridlines: false,
                    materializeDefaultGridlineStyle: false));
                plotArea.Append(CreateSharedValueAxis(plotArea, yAxisId, xAxisId,
                    "l", secondary: false,
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

            foreach (IGrouping<(OfficeChartType Kind, OfficeChartAxisGroup AxisGroup), SharedSeriesDescriptor> group in
                     descriptors.GroupBy(item => (item.Kind, item.AxisGroup)).OrderBy(item => ChartLayer(item.Key.Kind))) {
                uint categoryId = group.Key.AxisGroup == OfficeChartAxisGroup.Primary
                    ? primaryCategoryId : secondaryCategoryId;
                uint valueId = group.Key.AxisGroup == OfficeChartAxisGroup.Primary
                    ? primaryValueId : secondaryValueId;
                List<SharedSeriesDescriptor> items = group.ToList();
                if (IsBarOrColumnKind(group.Key.Kind)) {
                    plotArea.Append(CreateSharedBarChart(plotArea, group.Key.Kind, items, data.Categories, categoryId, valueId));
                } else if (IsLineKind(group.Key.Kind)) {
                    plotArea.Append(CreateSharedLineChart(plotArea, group.Key.Kind, items, data.Categories, categoryId, valueId));
                } else if (IsAreaKind(group.Key.Kind)) {
                    plotArea.Append(CreateSharedAreaChart(plotArea, group.Key.Kind, items, data.Categories, categoryId, valueId));
                } else if (group.Key.Kind == OfficeChartType.Radar) {
                    plotArea.Append(CreateSharedRadarChart(plotArea, items, data.Categories, categoryId, valueId));
                } else {
                    throw new NotSupportedException("Chart kind " + group.Key.Kind + " is not supported in this chart composition.");
                }
            }

            if (hasPrimary) {
                string categoryPosition = horizontal ? "l" : "b";
                string valuePosition = horizontal ? "b" : "l";
                plotArea.Append(CreateSharedCategoryAxis(plotArea, primaryCategoryId, primaryValueId,
                    categoryPosition, secondary: false));
                plotArea.Append(CreateSharedValueAxis(plotArea, primaryValueId, primaryCategoryId,
                    valuePosition, secondary: false,
                    showMajorGridlines: true,
                    materializeDefaultGridlineStyle: false));
            }
            if (hasSecondary) {
                plotArea.Append(CreateSharedCategoryAxis(plotArea, secondaryCategoryId, secondaryValueId,
                    "t", secondary: true));
                plotArea.Append(CreateSharedValueAxis(plotArea, secondaryValueId, secondaryCategoryId,
                    "r", secondary: true,
                    showMajorGridlines: false,
                    materializeDefaultGridlineStyle: false));
            }
        }

        private static OfficeXmlNode CreateSharedBarChart(OfficeXmlNode context, OfficeChartType kind,
            IReadOnlyList<SharedSeriesDescriptor> descriptors, IReadOnlyList<string> categories,
            uint categoryAxisId, uint valueAxisId) {
            string grouping = GetBarGrouping(kind);
            OfficeXmlNode chart = context.CreateElement(NsC, "barChart");
            chart.AppendVal(NsC, "barDir", IsHorizontalBarKind(kind) ? "bar" : "col");
            chart.AppendVal(NsC, "grouping", grouping);
            chart.AppendVal(NsC, "varyColors", Text(false));
            foreach (SharedSeriesDescriptor descriptor in descriptors) {
                chart.Append(CreateBarChartSeries(chart, descriptor.Index,
                    descriptor.Series, categories));
            }
            chart.Append(CreateDefaultDataLabels(chart));
            chart.AppendVal(NsC, "gapWidth",
                ((ushort)(uint)OfficeTheme.Default.ChartStyle.BarGapPercent).ToString(CultureInfo.InvariantCulture));
            chart.AppendVal(NsC, "overlap",
                ((sbyte)(grouping == "clustered" ? 0 : 100)).ToString(CultureInfo.InvariantCulture));
            chart.AppendVal(NsC, "axId", Text(categoryAxisId));
            chart.AppendVal(NsC, "axId", Text(valueAxisId));
            return chart;
        }

        private static OfficeXmlNode CreateSharedLineChart(OfficeXmlNode context, OfficeChartType kind,
            IReadOnlyList<SharedSeriesDescriptor> descriptors, IReadOnlyList<string> categories,
            uint categoryAxisId, uint valueAxisId) {
            OfficeXmlNode chart = context.CreateElement(NsC, "lineChart");
            chart.AppendVal(NsC, "grouping", GetLineGrouping(kind));
            chart.AppendVal(NsC, "varyColors", Text(false));
            foreach (SharedSeriesDescriptor descriptor in descriptors) {
                chart.Append(CreateLineChartSeries(chart, descriptor.Index,
                    descriptor.Series, categories));
            }
            chart.Append(CreateDefaultDataLabels(chart));
            chart.AppendVal(NsC, "axId", Text(categoryAxisId));
            chart.AppendVal(NsC, "axId", Text(valueAxisId));
            return chart;
        }

        private static OfficeXmlNode CreateSharedAreaChart(OfficeXmlNode context, OfficeChartType kind,
            IReadOnlyList<SharedSeriesDescriptor> descriptors, IReadOnlyList<string> categories,
            uint categoryAxisId, uint valueAxisId) {
            OfficeXmlNode chart = context.CreateElement(NsC, "areaChart");
            chart.AppendVal(NsC, "grouping", GetAreaGrouping(kind));
            chart.AppendVal(NsC, "varyColors", Text(false));
            foreach (SharedSeriesDescriptor descriptor in descriptors) {
                chart.Append(CreateSharedAreaSeries(chart, descriptor, categories));
            }
            chart.Append(CreateDefaultDataLabels(chart));
            chart.AppendVal(NsC, "axId", Text(categoryAxisId));
            chart.AppendVal(NsC, "axId", Text(valueAxisId));
            return chart;
        }

        private static OfficeXmlNode CreateSharedRadarChart(OfficeXmlNode context, IReadOnlyList<SharedSeriesDescriptor> descriptors,
            IReadOnlyList<string> categories, uint categoryAxisId, uint valueAxisId) {
            OfficeXmlNode chart = context.CreateElement(NsC, "radarChart");
            chart.AppendVal(NsC, "radarStyle", "marker");
            chart.AppendVal(NsC, "varyColors", Text(false));
            foreach (SharedSeriesDescriptor descriptor in descriptors) {
                chart.Append(CreateSharedRadarSeries(chart, descriptor, categories));
            }
            chart.Append(CreateDefaultDataLabels(chart));
            chart.AppendVal(NsC, "axId", Text(categoryAxisId));
            chart.AppendVal(NsC, "axId", Text(valueAxisId));
            return chart;
        }

        private static OfficeXmlNode CreateSharedAreaSeries(OfficeXmlNode context, SharedSeriesDescriptor descriptor,
            IReadOnlyList<string> categories) => CreateSharedCategorySeries(context, descriptor, categories);

        private static OfficeXmlNode CreateSharedRadarSeries(OfficeXmlNode context, SharedSeriesDescriptor descriptor,
            IReadOnlyList<string> categories) => CreateSharedCategorySeries(context, descriptor, categories);

        // The area and radar series carry the same children: idx, order, tx, cat and val.
        private static OfficeXmlNode CreateSharedCategorySeries(OfficeXmlNode context, SharedSeriesDescriptor descriptor,
            IReadOnlyList<string> categories) {
            string column = ColumnLetter(descriptor.Index + 2);
            int lastRow = categories.Count + 1;
            OfficeXmlNode series = context.CreateElement(NsC, "ser");
            series.AppendVal(NsC, "idx", Text((uint)descriptor.Index));
            series.AppendVal(NsC, "order", Text((uint)descriptor.Index));
            OfficeXmlNode text = series.AppendNew(NsC, "tx");
            text.Append(CreateStringReference(text, "Sheet1!$" + column + "$1",
                new[] { descriptor.Series.Name }));
            OfficeXmlNode categoryData = series.AppendNew(NsC, "cat");
            categoryData.Append(CreateStringReference(categoryData, "Sheet1!$A$2:$A$" + lastRow, categories));
            OfficeXmlNode values = series.AppendNew(NsC, "val");
            values.Append(CreateNumberReference(values, "Sheet1!$" + column + "$2:$" + column + "$" + lastRow,
                descriptor.Series.Values));
            return series;
        }

        private static OfficeXmlNode CreateSharedCategoryAxis(OfficeXmlNode context, uint axisId, uint crossingAxisId,
            string position, bool secondary) {
            OfficeXmlNode axis = context.CreateElement(NsC, "catAx");
            axis.AppendVal(NsC, "axId", Text(axisId));
            axis.AppendNew(NsC, "scaling").AppendVal(NsC, "orientation", "minMax");
            axis.AppendVal(NsC, "delete", Text(secondary));
            axis.AppendVal(NsC, "axPos", position);
            AppendGeneralNumberFormat(axis);
            axis.AppendVal(NsC, "majorTickMark", "none");
            axis.AppendVal(NsC, "minorTickMark", "none");
            axis.AppendVal(NsC, "tickLblPos", secondary ? "none" : "nextTo");
            axis.AppendVal(NsC, "crossAx", Text(crossingAxisId));
            axis.AppendVal(NsC, "crosses", secondary ? "max" : "autoZero");
            axis.AppendVal(NsC, "auto", Text(true));
            axis.AppendVal(NsC, "lblAlgn", "ctr");
            axis.AppendVal(NsC, "lblOffset", "100");
            axis.AppendVal(NsC, "noMultiLvlLbl", Text(false));
            return axis;
        }

        private static void AppendGeneralNumberFormat(OfficeXmlNode axis) {
            OfficeXmlNode format = axis.AppendNew(NsC, "numFmt");
            format.SetAttribute("formatCode", "General");
            format.SetAttribute("sourceLinked", Text(false));
        }

        private static OfficeXmlNode CreateSharedValueAxis(OfficeXmlNode context, uint axisId, uint crossingAxisId,
            string position, bool secondary,
            bool showMajorGridlines,
            bool materializeDefaultGridlineStyle) {
            OfficeXmlNode axis = context.CreateElement(NsC, "valAx");
            axis.AppendVal(NsC, "axId", Text(axisId));
            axis.AppendNew(NsC, "scaling").AppendVal(NsC, "orientation", "minMax");
            axis.AppendVal(NsC, "delete", Text(false));
            axis.AppendVal(NsC, "axPos", position);
            if (showMajorGridlines) {
                OfficeXmlNode gridlines = axis.AppendNew(NsC, "majorGridlines");
                if (materializeDefaultGridlineStyle) {
                    OfficeChartStyle style = OfficeChartStyle.Default;
                    OfficeXmlNode outline = gridlines.AppendNew(NsC, "spPr").AppendNew(NsA, "ln");
                    outline.SetAttribute("w", Text(checked((int)FromPoints(
                        style.GridLineWidth ?? 0.5D))));
                    outline.Append(CreateSharedSolidFill(outline, style.GridLineColor));
                }
            }
            AppendGeneralNumberFormat(axis);
            axis.AppendVal(NsC, "majorTickMark", "none");
            axis.AppendVal(NsC, "minorTickMark", "none");
            axis.AppendVal(NsC, "tickLblPos", "nextTo");
            axis.AppendVal(NsC, "crossAx", Text(crossingAxisId));
            axis.AppendVal(NsC, "crosses", secondary ? "max" : "autoZero");
            axis.AppendVal(NsC, "crossBetween", "between");
            return axis;
        }

        private static IEnumerable<uint> GetHiddenSharedLegendIndexes(OfficeChartData data, OfficeChartType kind) {
            kind = data.Series.FirstOrDefault()?.RenderKind ?? kind;
            if (kind == OfficeChartType.Pie || kind == OfficeChartType.Doughnut) {
                if (data.Series.Any(series => !series.ShowInLegend))
                    for (int point = 0; point < data.Categories.Count; point++) yield return (uint)point;
            } else {
                for (int index = 0; index < data.Series.Count; index++)
                    if (!data.Series[index].ShowInLegend) yield return (uint)index;
            }
        }

        internal static IEnumerable<OfficeXmlNode> GetSharedNativeChartLayers(OfficeXmlNode plotArea) {
            for (OfficeXmlNode element = plotArea.FirstChild; !element.IsNone; element = element.NextSibling) {
                if (element.LocalName.EndsWith("Chart", StringComparison.OrdinalIgnoreCase)) yield return element;
            }
        }

        private static bool IsOnlySharedScatterPlot(OfficeXmlNode plotArea) {
            if (plotArea.Child(NsC, "scatterChart").IsNone ||
                !GetSharedNativeChartLayers(plotArea).All(element => element.Is(NsC, "scatterChart"))) return false;
            var axisCounts = plotArea.Children(NsC, "valAx").ToList()
                .Where(axis => axis.Child(NsC, "axId").Val() != null)
                .GroupBy(axis => axis.Child(NsC, "axId").UIntVal()!.Value).ToDictionary(group => group.Key, group => group.Count());
            uint[]? expected = null;
            foreach (OfficeXmlNode layer in plotArea.Children(NsC, "scatterChart")) {
                var references = layer.Children(NsC, "axId").ToList().Take(3).ToList();
                if (references.Count != 2 || references.Any(axis => axis.Val() == null))
                    throw new NotSupportedException("In-place scatter updates require two referenced value axes per layer.");
                uint[] pair = references.Select(axis => axis.UIntVal()!.Value).ToArray();
                if (pair[0] == pair[1] || pair.Any(id => !axisCounts.TryGetValue(id, out int count) || count != 1) ||
                    expected != null && !expected.SequenceEqual(pair))
                    throw new NotSupportedException("In-place scatter updates require all layers to share the same ordered value-axis pair.");
                expected ??= pair;
            }
            return true;
        }

        private static OfficeXmlNode CreateSharedLegend(OfficeXmlNode context, OfficeChartData data, OfficeChartType kind) {
            OfficeXmlNode legend = context.CreateElement(NsC, "legend");
            legend.AppendVal(NsC, "legendPos", "b");
            foreach (uint index in GetHiddenSharedLegendIndexes(data, kind)) {
                OfficeXmlNode entry = legend.AppendNew(NsC, "legendEntry");
                entry.AppendVal(NsC, "idx", Text(index));
                entry.AppendVal(NsC, "delete", Text(true));
            }
            legend.AppendNew(NsC, "layout");
            legend.AppendVal(NsC, "overlay", Text(false));
            return legend;
        }

        private static void UpdateSharedLegend(OfficeXmlNode chart, OfficeChartData data, OfficeChartType kind,
            bool previousLegendUsesCategories) {
            OfficeXmlNode current = chart.Child(NsC, "legend");
            if (current.IsNone) {
                return;
            }

            OfficeXmlNode replacement = current.Clone();
            kind = data.Series.FirstOrDefault()?.RenderKind ?? kind;
            bool legendUsesCategories = kind is OfficeChartType.Pie or OfficeChartType.Doughnut;
            int entryCount = legendUsesCategories ? data.Categories.Count : data.Series.Count;
            var retained = (previousLegendUsesCategories == legendUsesCategories
                ? replacement.Children(NsC, "legendEntry").ToList() : new List<OfficeXmlNode>())
                .Where(entry => entry.Child(NsC, "idx").UIntVal() < entryCount)
                .Where(entry => HasOtherThanIndexAndDelete(entry))
                .GroupBy(entry => entry.Child(NsC, "idx").UIntVal()!.Value)
                .ToDictionary(group => group.Key, group => group.First().Clone());
            var hidden = new HashSet<uint>(GetHiddenSharedLegendIndexes(data, kind));
            replacement.RemoveChildren(NsC, "legendEntry");
            OfficeXmlNode insertBefore = default;
            foreach (OfficeXmlNode child in replacement.Children()) {
                if (!child.Is(NsC, "legendPos") && !child.Is(NsC, "legendEntry")) {
                    insertBefore = child;
                    break;
                }
            }
            foreach (uint index in retained.Keys.Concat(hidden).Distinct().OrderBy(value => value)) {
                OfficeXmlNode entry;
                if (!retained.TryGetValue(index, out entry)) {
                    entry = replacement.CreateElement(NsC, "legendEntry");
                    entry.AppendVal(NsC, "idx", Text(index));
                }
                OfficeXmlNode delete = entry.Child(NsC, "delete");
                if (!delete.IsNone) delete.Remove();
                if (hidden.Contains(index)) SetOrdered(entry, entry.CreateVal(NsC, "delete", Text(true)), CLegendEntryView.Order);
                replacement.InsertBefore(entry, insertBefore);
            }
            current.ReplaceWith(replacement);
        }

        private static bool HasOtherThanIndexAndDelete(OfficeXmlNode entry) {
            foreach (OfficeXmlNode child in entry.Children()) {
                if (!child.Is(NsC, "idx") && !child.Is(NsC, "delete")) return true;
            }
            return false;
        }

        private static IEnumerable<OfficeXmlNode> EnumerateSharedSeriesElements(OfficeXmlNode plotArea) {
            for (OfficeXmlNode chart = plotArea.FirstChild; !chart.IsNone; chart = chart.NextSibling) {
                if (chart.NamespaceId != NsC) continue;
                switch (chart.LocalName) {
                    case "barChart":
                    case "lineChart":
                    case "areaChart":
                    case "radarChart":
                    case "scatterChart":
                    case "bubbleChart":
                    case "pieChart":
                    case "doughnutChart":
                        for (OfficeXmlNode series = chart.Child(NsC, "ser"); !series.IsNone; series = series.NextSiblingNamed(NsC, "ser"))
                            yield return series;
                        break;
                }
            }
        }

        /// <summary>The child order of a <c>c:ser</c> under a chart layer with this local name, or null for another layer.</summary>
        private static OfficeXmlChildOrder? SeriesOrder(string? layer) => layer switch {
            "barChart" or "bar3DChart" => CBarSerView.Order,
            "lineChart" or "line3DChart" or "stockChart" => CLineSerView.Order,
            "scatterChart" => CScatterSerView.Order,
            "areaChart" or "area3DChart" => CAreaSerView.Order,
            "radarChart" => CRadarSerView.Order,
            "pieChart" or "pie3DChart" or "doughnutChart" or "ofPieChart" => CPieSerView.Order,
            "bubbleChart" => CBubbleSerView.Order,
            "surfaceChart" or "surface3DChart" => CSurfaceSerView.Order,
            _ => null
        };

        /// <summary>The layer a chart kind writes its series under, for a series that isn't attached to one yet.</summary>
        private static string? LayerOfKind(OfficeChartType? kind) => kind switch {
            null => null,
            OfficeChartType value when IsBarOrColumnKind(value) => "barChart",
            OfficeChartType value when IsLineKind(value) => "lineChart",
            OfficeChartType value when IsAreaKind(value) => "areaChart",
            OfficeChartType.Scatter => "scatterChart",
            OfficeChartType.Radar => "radarChart",
            OfficeChartType.Pie or OfficeChartType.Doughnut => "pieChart",
            OfficeChartType.Bubble => "bubbleChart",
            _ => null
        };

        /// <summary>
        /// The layer that holds a series: its parent's local name, or for a detached series the layer of
        /// <paramref name="kind"/>, which stands in for the SDK series class the caller built.
        /// </summary>
        private static string? SeriesLayer(OfficeXmlNode series, OfficeChartType? kind) {
            OfficeXmlNode parent = series.Parent;
            if (!parent.IsNone) return parent.NamespaceId == NsC ? parent.LocalName : null;
            return LayerOfKind(kind);
        }

        private static List<SharedSeriesDescriptor> DescribeSharedSeries(OfficeChartData data,
            OfficeChartType defaultKind) => data.Series.Select((series, index) =>
                new SharedSeriesDescriptor(index, series, series.RenderKind ?? defaultKind)).ToList();

        private static IReadOnlyList<double> ParseScatterCategories(IReadOnlyList<string> categories) {
            var values = new List<double>(categories.Count);
            foreach (string category in categories) {
                if (!OfficeIMO.Core.Internal.OfficeRoundTripText.TryParse(category, NumberStyles.Float, CultureInfo.InvariantCulture, out double value) ||
                    double.IsNaN(value) || double.IsInfinity(value)) {
                    throw new ArgumentException(
                        "Scatter chart categories must be finite invariant numeric values when a series has no XValues.",
                        nameof(categories));
                }
                values.Add(value);
            }
            return values;
        }

        private static bool IsBarOrColumnKind(OfficeChartType kind) =>
            kind == OfficeChartType.ColumnClustered || kind == OfficeChartType.ColumnStacked ||
            kind == OfficeChartType.ColumnStacked100 || kind == OfficeChartType.BarClustered ||
            kind == OfficeChartType.BarStacked || kind == OfficeChartType.BarStacked100;

        private static bool IsHorizontalBarKind(OfficeChartType kind) =>
            kind == OfficeChartType.BarClustered || kind == OfficeChartType.BarStacked ||
            kind == OfficeChartType.BarStacked100;

        private static bool IsLineKind(OfficeChartType kind) =>
            kind == OfficeChartType.Line || kind == OfficeChartType.LineStacked ||
            kind == OfficeChartType.LineStacked100;

        private static bool IsAreaKind(OfficeChartType kind) =>
            kind == OfficeChartType.Area || kind == OfficeChartType.AreaStacked ||
            kind == OfficeChartType.AreaStacked100;

        private static bool IsFilledSharedKind(OfficeChartType kind) =>
            IsBarOrColumnKind(kind) || IsAreaKind(kind) || kind == OfficeChartType.Pie ||
            kind == OfficeChartType.Doughnut || kind == OfficeChartType.Bubble;

        private static bool IsMarkerKind(OfficeChartType kind) => IsLineKind(kind) ||
            kind == OfficeChartType.Scatter || kind == OfficeChartType.Radar;

        private static string GetBarGrouping(OfficeChartType kind) {
            if (kind == OfficeChartType.ColumnStacked100 || kind == OfficeChartType.BarStacked100)
                return "percentStacked";
            if (kind == OfficeChartType.ColumnStacked || kind == OfficeChartType.BarStacked)
                return "stacked";
            return "clustered";
        }

        private static string GetLineGrouping(OfficeChartType kind) {
            if (kind == OfficeChartType.LineStacked100) return "percentStacked";
            if (kind == OfficeChartType.LineStacked) return "stacked";
            return "standard";
        }

        private static string GetAreaGrouping(OfficeChartType kind) {
            if (kind == OfficeChartType.AreaStacked100) return "percentStacked";
            if (kind == OfficeChartType.AreaStacked) return "stacked";
            return "standard";
        }

        private static int ChartLayer(OfficeChartType kind) {
            if (IsAreaKind(kind)) return 0;
            if (IsBarOrColumnKind(kind)) return 1;
            return 2;
        }

        private static string MapMarker(OfficeChartMarkerShape? marker) {
            if (!marker.HasValue) return "circle";
            switch (marker.Value) {
                case OfficeChartMarkerShape.Square: return "square";
                case OfficeChartMarkerShape.Diamond: return "diamond";
                case OfficeChartMarkerShape.Triangle: return "triangle";
                case OfficeChartMarkerShape.Dash: return "dash";
                case OfficeChartMarkerShape.Dot: return "dot";
                case OfficeChartMarkerShape.Plus: return "plus";
                case OfficeChartMarkerShape.X: return "x";
                case OfficeChartMarkerShape.Star: return "star";
                default: return "circle";
            }
        }

        private static string MapDash(OfficeStrokeDashStyle dash) {
            switch (dash) {
                case OfficeStrokeDashStyle.Dash: return "dash";
                case OfficeStrokeDashStyle.Dot: return "dot";
                case OfficeStrokeDashStyle.DashDot: return "dashDot";
                case OfficeStrokeDashStyle.DashDotDot: return "lgDashDotDot";
                default: return "solid";
            }
        }

        // SDK callers

        internal static void PopulateSharedChart(ChartPart chartPart, string embeddedRelId, OfficeChartData data,
            OfficeChartType defaultKind, OfficeTheme? theme = null) {
            ArgumentNullException.ThrowIfNull(chartPart);
            ValidateSharedChartData(data, defaultKind);
            // A fresh root with the declarations the chart part has always been written with.
            var chartSpace = new C.ChartSpace();
            chartSpace.AddNamespaceDeclaration("c", ChartNamespace);
            chartSpace.AddNamespaceDeclaration("a", DrawingNamespace);
            chartSpace.AddNamespaceDeclaration("r", RelationshipNamespace);
            chartPart.ChartSpace = chartSpace;
            PopulateSharedChartCore(new OfficeSdkChartPart(chartPart), embeddedRelId, data, defaultKind, theme);
        }

        internal static void UpdateSharedChartData(ChartPart chartPart, OfficeChartData data,
            OfficeChartType defaultKind, OfficeTheme? theme = null) {
            ArgumentNullException.ThrowIfNull(chartPart);
            UpdateSharedChartData(new OfficeSdkChartPart(chartPart), data, defaultKind, theme);
        }

        internal static IEnumerable<OpenXmlElement> GetSharedNativeChartLayers(C.PlotArea plotArea) =>
            GetSharedNativeChartLayers(OfficeXmlNode.FromSdk(plotArea)).Select(layer => (OpenXmlElement)layer.SdkElement!);
    }
}
