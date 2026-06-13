using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using DocumentFormat.OpenXml;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.Word {
    public partial class WordChart {
        /// <summary>
        /// Tries to create a dependency-free chart snapshot from cached Word chart data.
        /// </summary>
        public bool TryGetSnapshot(out WordChartSnapshot snapshot) {
            try {
                C.Chart? chart = _chartPart?.ChartSpace?.GetFirstChild<C.Chart>() ?? _chart;
                C.PlotArea? plotArea = chart?.GetFirstChild<C.PlotArea>();
                if (chart == null || plotArea == null) {
                    snapshot = null!;
                    return false;
                }

                Dictionary<A.SchemeColorValues, OfficeIMO.Drawing.OfficeColor> themeColors = GetThemeColors();

                if (CountSupportedChartElements(plotArea) > 1) {
                    snapshot = null!;
                    return false;
                }

                if (plotArea.GetFirstChild<C.BarChart>() is C.BarChart barChart) {
                    WordChartData? data = ReadCategorySeriesData(barChart.Elements<C.BarChartSeries>().Cast<OpenXmlCompositeElement>(), themeColors);
                    if (data == null) {
                        snapshot = null!;
                        return false;
                    }

                    snapshot = CreateSnapshot(chart, GetBarChartSnapshotKind(barChart), data);
                    return true;
                }

                if (plotArea.GetFirstChild<C.Bar3DChart>() is C.Bar3DChart bar3DChart) {
                    WordChartData? data = ReadCategorySeriesData(bar3DChart.Elements<C.BarChartSeries>().Cast<OpenXmlCompositeElement>(), themeColors);
                    if (data == null) {
                        snapshot = null!;
                        return false;
                    }

                    snapshot = CreateSnapshot(chart, GetBar3DChartSnapshotKind(bar3DChart), data);
                    return true;
                }

                if (plotArea.GetFirstChild<C.LineChart>() is C.LineChart lineChart) {
                    WordChartData? data = ReadCategorySeriesData(lineChart.Elements<C.LineChartSeries>().Cast<OpenXmlCompositeElement>(), themeColors);
                    if (data == null) {
                        snapshot = null!;
                        return false;
                    }

                    snapshot = CreateSnapshot(chart, GetLineChartSnapshotKind(lineChart), data);
                    return true;
                }

                if (plotArea.GetFirstChild<C.Line3DChart>() is C.Line3DChart line3DChart) {
                    WordChartData? data = ReadCategorySeriesData(line3DChart.Elements<C.LineChartSeries>().Cast<OpenXmlCompositeElement>(), themeColors);
                    if (data == null) {
                        snapshot = null!;
                        return false;
                    }

                    snapshot = CreateSnapshot(chart, WordChartSnapshotKind.Line, data);
                    return true;
                }

                if (plotArea.GetFirstChild<C.AreaChart>() is C.AreaChart areaChart) {
                    WordChartData? data = ReadCategorySeriesData(areaChart.Elements<C.AreaChartSeries>().Cast<OpenXmlCompositeElement>(), themeColors);
                    if (data == null) {
                        snapshot = null!;
                        return false;
                    }

                    snapshot = CreateSnapshot(chart, GetAreaChartSnapshotKind(areaChart), data);
                    return true;
                }

                if (plotArea.GetFirstChild<C.Area3DChart>() is C.Area3DChart area3DChart) {
                    WordChartData? data = ReadCategorySeriesData(area3DChart.Elements<C.AreaChartSeries>().Cast<OpenXmlCompositeElement>(), themeColors);
                    if (data == null) {
                        snapshot = null!;
                        return false;
                    }

                    snapshot = CreateSnapshot(chart, GetArea3DChartSnapshotKind(area3DChart), data);
                    return true;
                }

                if (plotArea.GetFirstChild<C.RadarChart>() is C.RadarChart radarChart) {
                    WordChartData? data = ReadCategorySeriesData(radarChart.Elements<C.RadarChartSeries>().Cast<OpenXmlCompositeElement>(), themeColors);
                    if (data == null) {
                        snapshot = null!;
                        return false;
                    }

                    snapshot = CreateSnapshot(chart, WordChartSnapshotKind.Radar, data);
                    return true;
                }

                if (plotArea.GetFirstChild<C.ScatterChart>() is C.ScatterChart scatterChart) {
                    WordChartData? data = ReadScatterSeriesData(scatterChart.Elements<C.ScatterChartSeries>(), themeColors);
                    if (data == null) {
                        snapshot = null!;
                        return false;
                    }

                    snapshot = CreateSnapshot(chart, WordChartSnapshotKind.Scatter, data);
                    return true;
                }

                if (plotArea.GetFirstChild<C.PieChart>() is C.PieChart pieChart) {
                    WordChartData? data = ReadCategorySeriesData(pieChart.Elements<C.PieChartSeries>().Cast<OpenXmlCompositeElement>(), themeColors);
                    if (data == null) {
                        snapshot = null!;
                        return false;
                    }

                    snapshot = CreateSnapshot(chart, WordChartSnapshotKind.Pie, data);
                    return true;
                }

                if (plotArea.GetFirstChild<C.Pie3DChart>() is C.Pie3DChart pie3DChart) {
                    WordChartData? data = ReadCategorySeriesData(pie3DChart.Elements<C.PieChartSeries>().Cast<OpenXmlCompositeElement>(), themeColors);
                    if (data == null) {
                        snapshot = null!;
                        return false;
                    }

                    snapshot = CreateSnapshot(chart, WordChartSnapshotKind.Pie, data);
                    return true;
                }

                if (plotArea.GetFirstChild<C.DoughnutChart>() is C.DoughnutChart doughnutChart) {
                    WordChartData? data = ReadCategorySeriesData(doughnutChart.Elements<C.PieChartSeries>().Cast<OpenXmlCompositeElement>(), themeColors);
                    if (data == null) {
                        snapshot = null!;
                        return false;
                    }

                    snapshot = CreateSnapshot(chart, WordChartSnapshotKind.Doughnut, data);
                    return true;
                }

                snapshot = null!;
                return false;
            } catch {
                snapshot = null!;
                return false;
            }
        }

        private static int CountSupportedChartElements(C.PlotArea plotArea) {
            return plotArea.Elements<C.BarChart>().Count()
                + plotArea.Elements<C.Bar3DChart>().Count()
                + plotArea.Elements<C.LineChart>().Count()
                + plotArea.Elements<C.Line3DChart>().Count()
                + plotArea.Elements<C.AreaChart>().Count()
                + plotArea.Elements<C.Area3DChart>().Count()
                + plotArea.Elements<C.RadarChart>().Count()
                + plotArea.Elements<C.ScatterChart>().Count()
                + plotArea.Elements<C.PieChart>().Count()
                + plotArea.Elements<C.Pie3DChart>().Count()
                + plotArea.Elements<C.DoughnutChart>().Count();
        }

        private WordChartSnapshot CreateSnapshot(C.Chart chart, WordChartSnapshotKind kind, WordChartData data) {
            return new WordChartSnapshot(
                ReadDrawingName(),
                ReadTitle(chart),
                kind,
                data,
                GetWidthPoints(),
                GetHeightPoints());
        }

        private static WordChartSnapshotKind GetBarChartSnapshotKind(C.BarChart chart) {
            C.BarDirectionValues direction = chart.GetFirstChild<C.BarDirection>()?.Val?.Value ?? C.BarDirectionValues.Column;
            C.BarGroupingValues grouping = chart.GetFirstChild<C.BarGrouping>()?.Val?.Value ?? C.BarGroupingValues.Clustered;
            return MapBarKind(direction, grouping);
        }

        private static WordChartSnapshotKind GetBar3DChartSnapshotKind(C.Bar3DChart chart) {
            C.BarDirectionValues direction = chart.GetFirstChild<C.BarDirection>()?.Val?.Value ?? C.BarDirectionValues.Column;
            C.BarGroupingValues grouping = chart.GetFirstChild<C.BarGrouping>()?.Val?.Value ?? C.BarGroupingValues.Clustered;
            return MapBarKind(direction, grouping);
        }

        private static WordChartSnapshotKind MapBarKind(C.BarDirectionValues direction, C.BarGroupingValues grouping) {
            bool horizontal = direction == C.BarDirectionValues.Bar;

            if (grouping == C.BarGroupingValues.Stacked) {
                return horizontal ? WordChartSnapshotKind.StackedBar : WordChartSnapshotKind.StackedColumn;
            }

            if (grouping == C.BarGroupingValues.PercentStacked) {
                return horizontal ? WordChartSnapshotKind.StackedBar100 : WordChartSnapshotKind.StackedColumn100;
            }

            return horizontal ? WordChartSnapshotKind.ClusteredBar : WordChartSnapshotKind.ClusteredColumn;
        }

        private static WordChartSnapshotKind GetLineChartSnapshotKind(C.LineChart chart) {
            C.GroupingValues grouping = chart.GetFirstChild<C.Grouping>()?.Val?.Value ?? C.GroupingValues.Standard;
            if (grouping == C.GroupingValues.Stacked) {
                return WordChartSnapshotKind.StackedLine;
            }

            if (grouping == C.GroupingValues.PercentStacked) {
                return WordChartSnapshotKind.StackedLine100;
            }

            return WordChartSnapshotKind.Line;
        }

        private static WordChartSnapshotKind GetAreaChartSnapshotKind(C.AreaChart chart) {
            C.GroupingValues grouping = chart.GetFirstChild<C.Grouping>()?.Val?.Value ?? C.GroupingValues.Standard;
            if (grouping == C.GroupingValues.Stacked) {
                return WordChartSnapshotKind.StackedArea;
            }

            if (grouping == C.GroupingValues.PercentStacked) {
                return WordChartSnapshotKind.StackedArea100;
            }

            return WordChartSnapshotKind.Area;
        }

        private static WordChartSnapshotKind GetArea3DChartSnapshotKind(C.Area3DChart chart) {
            C.GroupingValues grouping = chart.GetFirstChild<C.Grouping>()?.Val?.Value ?? C.GroupingValues.Standard;
            if (grouping == C.GroupingValues.Stacked) {
                return WordChartSnapshotKind.StackedArea;
            }

            if (grouping == C.GroupingValues.PercentStacked) {
                return WordChartSnapshotKind.StackedArea100;
            }

            return WordChartSnapshotKind.Area;
        }

        private static WordChartData? ReadCategorySeriesData(IEnumerable<OpenXmlCompositeElement> seriesElements, IReadOnlyDictionary<A.SchemeColorValues, OfficeIMO.Drawing.OfficeColor> themeColors) {
            var seriesList = seriesElements.ToList();
            if (seriesList.Count == 0) {
                return null;
            }

            IReadOnlyList<string> categories = Array.Empty<string>();
            IReadOnlyList<string> fallbackCategories = Array.Empty<string>();
            for (int i = 0; i < seriesList.Count; i++) {
                IReadOnlyList<double> values = ReadCachedNumbers(seriesList[i].GetFirstChild<C.Values>());
                if (values.Count == 0) {
                    continue;
                }

                categories = ReadCachedStrings(seriesList[i].GetFirstChild<C.CategoryAxisData>());
                if (categories.Count > 0) {
                    break;
                }

                if (fallbackCategories.Count == 0) {
                    fallbackCategories = CreateFallbackCategories(values.Count);
                }
            }

            if (categories.Count == 0) {
                categories = fallbackCategories;
                if (categories.Count == 0) {
                    return null;
                }
            }

            var series = new List<WordChartSeries>();
            for (int i = 0; i < seriesList.Count; i++) {
                OpenXmlCompositeElement seriesElement = seriesList[i];
                IReadOnlyList<double> values = NormalizeValues(ReadCachedNumbers(seriesElement.GetFirstChild<C.Values>()), categories.Count);
                if (values.Count == 0) {
                    continue;
                }

                string name = ReadSeriesName(seriesElement);
                if (string.IsNullOrWhiteSpace(name)) {
                    name = "Series " + (i + 1).ToString(CultureInfo.InvariantCulture);
                }

                series.Add(new WordChartSeries(
                    name,
                    values,
                    color: ReadSeriesColor(seriesElement, themeColors),
                    pointColors: ReadPointColors(seriesElement, values.Count, themeColors)));
            }

            return series.Count == 0 ? null : new WordChartData(categories, series);
        }

        private static WordChartData? ReadScatterSeriesData(IEnumerable<C.ScatterChartSeries> seriesElements, IReadOnlyDictionary<A.SchemeColorValues, OfficeIMO.Drawing.OfficeColor> themeColors) {
            var seriesList = seriesElements.ToList();
            if (seriesList.Count == 0) {
                return null;
            }

            var series = new List<WordChartSeries>();
            IReadOnlyList<double>? categoryXValues = null;
            for (int i = 0; i < seriesList.Count; i++) {
                C.ScatterChartSeries seriesElement = seriesList[i];
                IReadOnlyList<double> xValues = ReadCachedNumbers(seriesElement.GetFirstChild<C.XValues>());
                IReadOnlyList<double> yValues = ReadCachedNumbers(seriesElement.GetFirstChild<C.YValues>());
                int pointCount = Math.Min(xValues.Count, yValues.Count);
                if (pointCount == 0) {
                    continue;
                }

                IReadOnlyList<double> values = NormalizeValues(yValues, pointCount);
                if (values.Count == 0) {
                    continue;
                }

                categoryXValues ??= xValues.Take(pointCount).ToList();
                string name = ReadSeriesName(seriesElement);
                if (string.IsNullOrWhiteSpace(name)) {
                    name = "Series " + (i + 1).ToString(CultureInfo.InvariantCulture);
                }

                series.Add(new WordChartSeries(
                    name,
                    values,
                    xValues.Take(pointCount).ToList(),
                    ReadSeriesColor(seriesElement, themeColors),
                    ReadPointColors(seriesElement, values.Count, themeColors)));
            }

            if (series.Count == 0 || categoryXValues == null || categoryXValues.Count == 0) {
                return null;
            }

            var categories = categoryXValues
                .Select(value => value.ToString(CultureInfo.InvariantCulture))
                .ToList();
            return new WordChartData(categories, series);
        }

        private static string? ReadTitle(C.Chart chart) {
            C.ChartText? chartText = chart.GetFirstChild<C.Title>()?.GetFirstChild<C.ChartText>();
            if (chartText == null) {
                return null;
            }

            string text = string.Concat(chartText.Descendants<A.Text>().Select(item => item.Text));
            if (!string.IsNullOrWhiteSpace(text)) {
                return text.Trim();
            }

            IReadOnlyList<string> cached = ReadCachedStrings(chartText);
            return cached.Count > 0 && !string.IsNullOrWhiteSpace(cached[0]) ? cached[0].Trim() : null;
        }

        private static string ReadSeriesName(OpenXmlElement seriesElement) {
            C.SeriesText? seriesText = seriesElement.GetFirstChild<C.SeriesText>();
            if (seriesText == null) {
                return string.Empty;
            }

            IReadOnlyList<string> cached = ReadCachedStrings(seriesText);
            if (cached.Count > 0) {
                return cached[0] ?? string.Empty;
            }

            string richText = string.Concat(seriesText.Descendants<A.Text>().Select(item => item.Text));
            return richText.Trim();
        }

        private static OfficeIMO.Drawing.OfficeColor? ReadSeriesColor(OpenXmlElement seriesElement, IReadOnlyDictionary<A.SchemeColorValues, OfficeIMO.Drawing.OfficeColor> themeColors) {
            C.ChartShapeProperties? shapeProperties = seriesElement.GetFirstChild<C.ChartShapeProperties>();
            return ReadShapeFillColor(shapeProperties, themeColors)
                ?? ReadShapeOutlineColor(shapeProperties, themeColors);
        }

        private static IReadOnlyList<OfficeIMO.Drawing.OfficeColor?>? ReadPointColors(OpenXmlElement seriesElement, int pointCount, IReadOnlyDictionary<A.SchemeColorValues, OfficeIMO.Drawing.OfficeColor> themeColors) {
            if (pointCount <= 0) {
                return null;
            }

            var colors = new OfficeIMO.Drawing.OfficeColor?[pointCount];
            bool hasColor = false;
            foreach (C.DataPoint dataPoint in seriesElement.Elements<C.DataPoint>()) {
                int index = (int)(dataPoint.GetFirstChild<C.Index>()?.Val?.Value ?? uint.MaxValue);
                if (index < 0 || index >= pointCount) {
                    continue;
                }

                OfficeIMO.Drawing.OfficeColor? color = ReadShapeFillColor(dataPoint.GetFirstChild<C.ChartShapeProperties>(), themeColors);
                if (color.HasValue) {
                    colors[index] = color.Value;
                    hasColor = true;
                }
            }

            return hasColor ? colors : null;
        }

        private static OfficeIMO.Drawing.OfficeColor? ReadShapeFillColor(C.ChartShapeProperties? shapeProperties, IReadOnlyDictionary<A.SchemeColorValues, OfficeIMO.Drawing.OfficeColor> themeColors) {
            A.SolidFill? fill = shapeProperties?.GetFirstChild<A.SolidFill>();
            return ReadSolidFillColor(fill, themeColors);
        }

        private static OfficeIMO.Drawing.OfficeColor? ReadShapeOutlineColor(C.ChartShapeProperties? shapeProperties, IReadOnlyDictionary<A.SchemeColorValues, OfficeIMO.Drawing.OfficeColor> themeColors) {
            A.SolidFill? fill = shapeProperties?.GetFirstChild<A.Outline>()?.GetFirstChild<A.SolidFill>();
            return ReadSolidFillColor(fill, themeColors);
        }

        private static OfficeIMO.Drawing.OfficeColor? ReadSolidFillColor(A.SolidFill? fill, IReadOnlyDictionary<A.SchemeColorValues, OfficeIMO.Drawing.OfficeColor> themeColors) {
            if (fill == null) {
                return null;
            }

            return ReadColorChoice(fill.GetFirstChild<A.RgbColorModelHex>(), fill.GetFirstChild<A.SchemeColor>(), fill.GetFirstChild<A.SystemColor>(), themeColors);
        }

        private Dictionary<A.SchemeColorValues, OfficeIMO.Drawing.OfficeColor> GetThemeColors() {
            var colors = new Dictionary<A.SchemeColorValues, OfficeIMO.Drawing.OfficeColor>();
            A.ColorScheme? colorScheme = _document.MainDocumentPartRoot.ThemePart?.Theme?.ThemeElements?.ColorScheme;
            if (colorScheme == null) {
                return colors;
            }

            AddThemeColor(colors, A.SchemeColorValues.Dark1, colorScheme.GetFirstChild<A.Dark1Color>());
            AddThemeColor(colors, A.SchemeColorValues.Light1, colorScheme.GetFirstChild<A.Light1Color>());
            AddThemeColor(colors, A.SchemeColorValues.Dark2, colorScheme.GetFirstChild<A.Dark2Color>());
            AddThemeColor(colors, A.SchemeColorValues.Light2, colorScheme.GetFirstChild<A.Light2Color>());
            AddThemeColor(colors, A.SchemeColorValues.Accent1, colorScheme.GetFirstChild<A.Accent1Color>());
            AddThemeColor(colors, A.SchemeColorValues.Accent2, colorScheme.GetFirstChild<A.Accent2Color>());
            AddThemeColor(colors, A.SchemeColorValues.Accent3, colorScheme.GetFirstChild<A.Accent3Color>());
            AddThemeColor(colors, A.SchemeColorValues.Accent4, colorScheme.GetFirstChild<A.Accent4Color>());
            AddThemeColor(colors, A.SchemeColorValues.Accent5, colorScheme.GetFirstChild<A.Accent5Color>());
            AddThemeColor(colors, A.SchemeColorValues.Accent6, colorScheme.GetFirstChild<A.Accent6Color>());
            AddThemeColor(colors, A.SchemeColorValues.Hyperlink, colorScheme.GetFirstChild<A.Hyperlink>());
            AddThemeColor(colors, A.SchemeColorValues.FollowedHyperlink, colorScheme.GetFirstChild<A.FollowedHyperlinkColor>());
            AddThemeAlias(colors, A.SchemeColorValues.Background1, A.SchemeColorValues.Light1);
            AddThemeAlias(colors, A.SchemeColorValues.Text1, A.SchemeColorValues.Dark1);
            AddThemeAlias(colors, A.SchemeColorValues.Background2, A.SchemeColorValues.Light2);
            AddThemeAlias(colors, A.SchemeColorValues.Text2, A.SchemeColorValues.Dark2);
            return colors;
        }

        private static void AddThemeAlias(
            Dictionary<A.SchemeColorValues, OfficeIMO.Drawing.OfficeColor> colors,
            A.SchemeColorValues alias,
            A.SchemeColorValues target) {
            if (!colors.ContainsKey(alias) && colors.TryGetValue(target, out var color)) {
                colors[alias] = color;
            }
        }

        private static void AddThemeColor(Dictionary<A.SchemeColorValues, OfficeIMO.Drawing.OfficeColor> colors, A.SchemeColorValues key, OpenXmlElement? element) {
            OfficeIMO.Drawing.OfficeColor? color = ReadColorChoice(
                element?.GetFirstChild<A.RgbColorModelHex>(),
                element?.GetFirstChild<A.SchemeColor>(),
                element?.GetFirstChild<A.SystemColor>(),
                colors);
            if (color.HasValue) {
                colors[key] = color.Value;
            }
        }

        private static OfficeIMO.Drawing.OfficeColor? ReadColorChoice(
            A.RgbColorModelHex? rgb,
            A.SchemeColor? scheme,
            A.SystemColor? system,
            IReadOnlyDictionary<A.SchemeColorValues, OfficeIMO.Drawing.OfficeColor> themeColors) {
            if (OfficeIMO.Drawing.OfficeColor.TryParseHex(rgb?.Val?.Value, out var directColor)) {
                return ApplyColorTransforms(directColor, rgb!);
            }

            if (scheme?.Val?.Value != null && themeColors.TryGetValue(scheme.Val.Value, out var themeColor)) {
                return ApplyColorTransforms(themeColor, scheme);
            }

            if (OfficeIMO.Drawing.OfficeColor.TryParseHex(system?.LastColor?.Value, out var systemColor)) {
                return ApplyColorTransforms(systemColor, system!);
            }

            return null;
        }

        private static OfficeIMO.Drawing.OfficeColor ApplyColorTransforms(OfficeIMO.Drawing.OfficeColor color, OpenXmlElement colorElement) {
            double red = color.R;
            double green = color.G;
            double blue = color.B;
            double alpha = color.A;

            foreach (OpenXmlElement transform in colorElement.ChildElements) {
                switch (transform) {
                    case A.LuminanceModulation luminanceModulation when luminanceModulation.Val != null:
                        double luminanceMultiplier = ClampPercentage(luminanceModulation.Val.Value);
                        red *= luminanceMultiplier;
                        green *= luminanceMultiplier;
                        blue *= luminanceMultiplier;
                        break;
                    case A.LuminanceOffset luminanceOffset when luminanceOffset.Val != null:
                        double luminanceOffsetValue = 255D * ClampPercentage(luminanceOffset.Val.Value);
                        red += luminanceOffsetValue;
                        green += luminanceOffsetValue;
                        blue += luminanceOffsetValue;
                        break;
                    case A.Tint tint when tint.Val != null:
                        double tintAmount = ClampPercentage(tint.Val.Value);
                        red = red + (255D - red) * tintAmount;
                        green = green + (255D - green) * tintAmount;
                        blue = blue + (255D - blue) * tintAmount;
                        break;
                    case A.Shade shade when shade.Val != null:
                        double shadeAmount = ClampPercentage(shade.Val.Value);
                        red *= shadeAmount;
                        green *= shadeAmount;
                        blue *= shadeAmount;
                        break;
                    case A.Alpha alphaTransform when alphaTransform.Val != null:
                        alpha = 255D * ClampPercentage(alphaTransform.Val.Value);
                        break;
                    case A.AlphaModulation alphaModulation when alphaModulation.Val != null:
                        alpha *= ClampPercentage(alphaModulation.Val.Value);
                        break;
                    case A.AlphaOffset alphaOffset when alphaOffset.Val != null:
                        alpha += 255D * ClampPercentage(alphaOffset.Val.Value);
                        break;
                }
            }

            return OfficeIMO.Drawing.OfficeColor.FromRgba(ToByte(red), ToByte(green), ToByte(blue), ToByte(alpha));
        }

        private static double ClampPercentage(int value) {
            if (value <= 0) {
                return 0D;
            }

            return value >= 100000 ? 1D : value / 100000D;
        }

        private static byte ToByte(double value) {
            if (value <= 0D) {
                return 0;
            }

            if (value >= 255D) {
                return 255;
            }

            return (byte)Math.Round(value, MidpointRounding.AwayFromZero);
        }

        private static IReadOnlyList<string> ReadCachedStrings(OpenXmlElement? container) {
            if (container == null) {
                return Array.Empty<string>();
            }

            List<C.StringPoint> stringPoints = container.Descendants<C.StringPoint>().OrderBy(point => point.Index?.Value ?? 0U).ToList();
            if (stringPoints.Count > 0) {
                return CreateIndexedCache(
                    container,
                    stringPoints,
                    point => point.Index?.Value,
                    point => point.NumericValue?.Text ?? string.Empty,
                    string.Empty);
            }

            List<C.NumericPoint> numericPoints = container.Descendants<C.NumericPoint>().OrderBy(point => point.Index?.Value ?? 0U).ToList();
            if (numericPoints.Count > 0) {
                return CreateIndexedCache(
                    container,
                    numericPoints,
                    point => point.Index?.Value,
                    point => point.NumericValue?.Text ?? string.Empty,
                    string.Empty);
            }

            return Array.Empty<string>();
        }

        private static IReadOnlyList<double> ReadCachedNumbers(OpenXmlElement? container) {
            if (container == null) {
                return Array.Empty<double>();
            }

            List<C.NumericPoint> points = container.Descendants<C.NumericPoint>().OrderBy(point => point.Index?.Value ?? 0U).ToList();
            if (points.Count == 0) {
                return Array.Empty<double>();
            }

            return CreateIndexedCache(
                container,
                points,
                point => point.Index?.Value,
                point => {
                    string? text = point.NumericValue?.Text;
                    if (double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out double value) &&
                        !double.IsNaN(value) &&
                        !double.IsInfinity(value)) {
                        return value;
                    }

                    return 0D;
                },
                0D);
        }

        private static IReadOnlyList<TValue> CreateIndexedCache<TPoint, TValue>(
            OpenXmlElement container,
            IReadOnlyList<TPoint> points,
            Func<TPoint, uint?> getIndex,
            Func<TPoint, TValue> getValue,
            TValue defaultValue) {
            int length = GetCachedPointLength(container, points, getIndex);
            var values = Enumerable.Repeat(defaultValue, length).ToArray();
            for (int i = 0; i < points.Count; i++) {
                TPoint point = points[i];
                uint? rawIndex = getIndex(point);
                int index = rawIndex.HasValue && rawIndex.Value <= int.MaxValue
                    ? (int)rawIndex.Value
                    : i;
                if (index >= 0 && index < values.Length) {
                    values[index] = getValue(point);
                }
            }

            return values;
        }

        private static int GetCachedPointLength<TPoint>(OpenXmlElement container, IReadOnlyList<TPoint> points, Func<TPoint, uint?> getIndex) {
            uint? pointCount = container.Descendants<C.PointCount>().FirstOrDefault()?.Val?.Value;
            uint maxIndex = 0U;
            bool hasIndexedPoint = false;
            for (int i = 0; i < points.Count; i++) {
                uint? index = getIndex(points[i]);
                if (!index.HasValue) {
                    continue;
                }

                hasIndexedPoint = true;
                if (index.Value > maxIndex) {
                    maxIndex = index.Value;
                }
            }

            uint indexedLength = hasIndexedPoint ? maxIndex + 1U : (uint)points.Count;
            uint length = Math.Max(pointCount ?? 0U, indexedLength);
            if (length > int.MaxValue) {
                return points.Count;
            }

            return (int)length;
        }

        private static IReadOnlyList<string> CreateFallbackCategories(int count) {
            if (count <= 0) {
                return Array.Empty<string>();
            }

            var categories = new List<string>(count);
            for (int i = 0; i < count; i++) {
                categories.Add("Category " + (i + 1).ToString(CultureInfo.InvariantCulture));
            }

            return categories;
        }

        private static IReadOnlyList<double> NormalizeValues(IReadOnlyList<double> values, int count) {
            if (count <= 0 || values.Count == 0) {
                return Array.Empty<double>();
            }

            var normalized = new double[count];
            int take = Math.Min(values.Count, count);
            for (int i = 0; i < take; i++) {
                normalized[i] = values[i];
            }

            return normalized;
        }

        private string ReadDrawingName() {
            return _drawing?.Inline?.DocProperties?.Name?.Value ?? string.Empty;
        }

        private double GetWidthPoints() {
            long emu = _drawing?.Inline?.Extent?.Cx?.Value ?? 0L;
            return emu > 0 ? EmuToPoints(emu) : 450D;
        }

        private double GetHeightPoints() {
            long emu = _drawing?.Inline?.Extent?.Cy?.Value ?? 0L;
            return emu > 0 ? EmuToPoints(emu) : 300D;
        }

        private static double EmuToPoints(long emu) {
            return emu * 72D / EnglishMetricUnitsPerInch;
        }
    }
}
