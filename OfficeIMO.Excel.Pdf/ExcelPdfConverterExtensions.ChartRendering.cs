using OfficeIMO.Drawing;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Excel.Pdf {
    public static partial class ExcelPdfConverterExtensions {
        private static void AddWorksheetChart(PdfCore.PdfItemCompose item, WorksheetChartExportData chart) {
            ExcelChartSnapshot snapshot = chart.Snapshot;
            string title = string.IsNullOrWhiteSpace(snapshot.Title) ? snapshot.Name : snapshot.Title!;
            if (!string.IsNullOrWhiteSpace(title)) {
                item.H2(title, PdfCore.PdfAlign.Left, PdfCore.PdfColor.FromRgb(31, 78, 121));
            }

            item.Drawing(CreateChartDrawing(snapshot), PdfCore.PdfAlign.Left, spacingBefore: 2, spacingAfter: 6);
            item.Table(CreateChartLegendRows(snapshot), PdfCore.PdfAlign.Left, CreateChartLegendStyle(GetChartLegendColorCount(snapshot)));
        }

        private static OfficeDrawing CreateChartDrawing(ExcelChartSnapshot snapshot) {
            double width = Math.Min(420D, Math.Max(240D, PixelsToPoints(snapshot.WidthPixels)));
            double height = Math.Min(260D, Math.Max(150D, PixelsToPoints(snapshot.HeightPixels)));
            var drawing = new OfficeDrawing(width, height);

            AddShape(drawing, OfficeShape.Rectangle(width, height), 0, 0, OfficeColor.FromRgb(250, 252, 255), OfficeColor.FromRgb(183, 194, 207), 0.75);

            if (IsPieChart(snapshot.ChartType) || IsDoughnutChart(snapshot.ChartType)) {
                AddPieSeries(drawing, snapshot, width, height, IsDoughnutChart(snapshot.ChartType));
                return drawing;
            }

            if (IsRadarChart(snapshot.ChartType)) {
                AddRadarSeries(drawing, snapshot, width, height);
                return drawing;
            }

            double plotLeft = 36D;
            double plotTop = 18D;
            double plotRight = 12D;
            double plotBottom = 28D;
            double plotWidth = Math.Max(20D, width - plotLeft - plotRight);
            double plotHeight = Math.Max(20D, height - plotTop - plotBottom);
            double plotBottomY = plotTop + plotHeight;

            AddShape(drawing, OfficeShape.Line(0, 0, plotWidth, 0), plotLeft, plotBottomY, null, OfficeColor.FromRgb(80, 90, 100), 0.75);
            AddShape(drawing, OfficeShape.Line(0, 0, 0, plotHeight), plotLeft, plotTop, null, OfficeColor.FromRgb(80, 90, 100), 0.75);
            for (int i = 1; i <= 3; i++) {
                double y = plotTop + (plotHeight * i / 4D);
                AddShape(drawing, OfficeShape.Line(0, 0, plotWidth, 0), plotLeft, y, null, OfficeColor.FromRgb(226, 232, 240), 0.5);
            }

            if (IsAreaChart(snapshot.ChartType)) {
                AddAreaSeries(drawing, snapshot, plotLeft, plotTop, plotWidth, plotHeight);
            } else if (IsScatterChart(snapshot.ChartType)) {
                AddScatterSeries(drawing, snapshot, plotLeft, plotTop, plotWidth, plotHeight);
            } else if (IsLineChart(snapshot.ChartType)) {
                AddLineSeries(drawing, snapshot, plotLeft, plotTop, plotWidth, plotHeight);
            } else {
                AddBarSeries(drawing, snapshot, plotLeft, plotTop, plotWidth, plotHeight);
            }

            return drawing;
        }

        private static void AddBarSeries(OfficeDrawing drawing, ExcelChartSnapshot snapshot, double plotLeft, double plotTop, double plotWidth, double plotHeight) {
            IReadOnlyList<string> categories = snapshot.Data.Categories;
            IReadOnlyList<ExcelChartSeries> series = snapshot.Data.Series;
            if (categories.Count == 0 || series.Count == 0) {
                return;
            }

            double slot = plotWidth / categories.Count;
            double groupWidth = slot * 0.68D;
            bool horizontal = IsBarChart(snapshot.ChartType);
            bool stacked = IsStackedBarOrColumnChart(snapshot.ChartType) || IsPercentStackedBarOrColumnChart(snapshot.ChartType);
            bool percentStacked = IsPercentStackedBarOrColumnChart(snapshot.ChartType);
            double barWidth = Math.Max(2D, stacked ? groupWidth : groupWidth / series.Count);
            (double min, double max) = percentStacked
                ? (0D, 1D)
                : stacked
                    ? GetStackedSeriesRange(series, categories.Count)
                    : GetFiniteSeriesRange(series);
            min = Math.Min(0D, min);
            max = Math.Max(0D, max);
            if (max <= min) {
                max = min + 1D;
            }

            for (int category = 0; category < categories.Count; category++) {
                double positiveBase = 0D;
                double negativeBase = 0D;
                double percentTotal = percentStacked ? GetPositiveCategoryTotal(series, category) : 0D;
                for (int s = 0; s < series.Count; s++) {
                    double value = GetSeriesValue(series[s], category);
                    if (value == 0D) {
                        continue;
                    }

                    double baseline = 0D;
                    double plottedValue = value;
                    if (stacked) {
                        if (percentStacked) {
                            plottedValue = percentTotal <= 0D ? 0D : Math.Max(0D, value) / percentTotal;
                        }

                        baseline = plottedValue >= 0D ? positiveBase : negativeBase;
                        if (plottedValue >= 0D) {
                            positiveBase += plottedValue;
                        } else {
                            negativeBase += plottedValue;
                        }
                    }

                    OfficeColor color = GetChartSeriesColor(s);
                    if (horizontal) {
                        double categoryHeight = plotHeight / categories.Count;
                        double rowHeight = Math.Max(2D, categoryHeight * 0.68D / (stacked ? 1D : series.Count));
                        double y = plotTop + (categoryHeight * category) + (categoryHeight * 0.16D) + (stacked ? 0D : rowHeight * s);
                        double x1 = ToPlotX(baseline, min, max, plotLeft, plotWidth);
                        double x2 = ToPlotX(stacked ? baseline + plottedValue : plottedValue, min, max, plotLeft, plotWidth);
                        double x = Math.Min(x1, x2);
                        double w = Math.Max(1D, Math.Abs(x2 - x1));
                        AddShape(drawing, OfficeShape.Rectangle(w, rowHeight), x, y, color, null, 0);
                    } else {
                        double x = plotLeft + (slot * category) + ((slot - groupWidth) / 2D) + (stacked ? 0D : barWidth * s);
                        double y1 = ToPlotY(baseline, min, max, plotTop, plotHeight);
                        double y2 = ToPlotY(stacked ? baseline + plottedValue : plottedValue, min, max, plotTop, plotHeight);
                        double y = Math.Min(y1, y2);
                        double h = Math.Max(1D, Math.Abs(y2 - y1));
                        AddShape(drawing, OfficeShape.Rectangle(barWidth * 0.88D, h), x, y, color, null, 0);
                    }
                }
            }
        }

        private static void AddAreaSeries(OfficeDrawing drawing, ExcelChartSnapshot snapshot, double plotLeft, double plotTop, double plotWidth, double plotHeight) {
            IReadOnlyList<string> categories = snapshot.Data.Categories;
            IReadOnlyList<ExcelChartSeries> series = snapshot.Data.Series;
            if (categories.Count < 2 || series.Count == 0) {
                return;
            }

            bool stacked = IsStackedAreaChart(snapshot.ChartType) || IsPercentStackedAreaChart(snapshot.ChartType);
            bool percentStacked = IsPercentStackedAreaChart(snapshot.ChartType);
            (double min, double max) = percentStacked
                ? (0D, 1D)
                : stacked
                    ? GetStackedSeriesRange(series, categories.Count)
                    : GetFiniteSeriesRange(series);
            double step = plotWidth / (categories.Count - 1);
            var positiveCumulative = new double[categories.Count];
            var negativeCumulative = new double[categories.Count];

            for (int s = 0; s < series.Count; s++) {
                OfficeColor color = GetChartSeriesColor(s);
                var topPoints = new List<OfficePoint>(categories.Count);
                var bottomPoints = new List<OfficePoint>(categories.Count);

                for (int i = 0; i < categories.Count; i++) {
                    double value = GetSeriesValue(series[s], i);
                    double rawValue = percentStacked ? Math.Max(0D, value) : value;
                    double baseline = stacked
                        ? (rawValue >= 0D ? positiveCumulative[i] : negativeCumulative[i])
                        : 0D;
                    double topValue = baseline + rawValue;

                    if (percentStacked) {
                        double total = GetPositiveCategoryTotal(series, i);
                        baseline = total <= 0D ? 0D : baseline / total;
                        topValue = total <= 0D ? 0D : topValue / total;
                    }

                    double x = plotLeft + step * i;
                    topPoints.Add(new OfficePoint(x, ToPlotY(topValue, min, max, plotTop, plotHeight)));
                    bottomPoints.Add(new OfficePoint(x, ToPlotY(baseline, min, max, plotTop, plotHeight)));
                }

                var areaPoints = new List<OfficePoint>(topPoints.Count + bottomPoints.Count);
                areaPoints.AddRange(topPoints);
                for (int i = bottomPoints.Count - 1; i >= 0; i--) {
                    areaPoints.Add(bottomPoints[i]);
                }

                AddPolygonShape(drawing, areaPoints, color, color, 0.5D, 0.32D);
                AddPointLine(drawing, topPoints, color, 1.4D);

                if (stacked) {
                    for (int i = 0; i < categories.Count; i++) {
                        double value = percentStacked ? Math.Max(0D, GetSeriesValue(series[s], i)) : GetSeriesValue(series[s], i);
                        if (value >= 0D) {
                            positiveCumulative[i] += value;
                        } else {
                            negativeCumulative[i] += value;
                        }
                    }
                }
            }
        }

        private static void AddLineSeries(OfficeDrawing drawing, ExcelChartSnapshot snapshot, double plotLeft, double plotTop, double plotWidth, double plotHeight) {
            IReadOnlyList<string> categories = snapshot.Data.Categories;
            IReadOnlyList<ExcelChartSeries> series = snapshot.Data.Series;
            if (categories.Count < 2 || series.Count == 0) {
                return;
            }

            (double min, double max) = GetFiniteSeriesRange(series);
            double step = plotWidth / (categories.Count - 1);
            for (int s = 0; s < series.Count; s++) {
                OfficeColor color = GetChartSeriesColor(s);
                for (int i = 1; i < categories.Count; i++) {
                    double x1 = plotLeft + step * (i - 1);
                    double y1 = ToPlotY(GetSeriesValue(series[s], i - 1), min, max, plotTop, plotHeight);
                    double x2 = plotLeft + step * i;
                    double y2 = ToPlotY(GetSeriesValue(series[s], i), min, max, plotTop, plotHeight);
                    double minX = Math.Min(x1, x2);
                    double minY = Math.Min(y1, y2);
                    AddShape(drawing, OfficeShape.Line(x1 - minX, y1 - minY, x2 - minX, y2 - minY), minX, minY, null, color, 1.75);
                }

                for (int i = 0; i < categories.Count; i++) {
                    double x = plotLeft + step * i - 2D;
                    double y = ToPlotY(GetSeriesValue(series[s], i), min, max, plotTop, plotHeight) - 2D;
                    AddShape(drawing, OfficeShape.Ellipse(4D, 4D), x, y, OfficeColor.White, color, 1D);
                }
            }
        }

        private static void AddScatterSeries(OfficeDrawing drawing, ExcelChartSnapshot snapshot, double plotLeft, double plotTop, double plotWidth, double plotHeight) {
            IReadOnlyList<string> categories = snapshot.Data.Categories;
            IReadOnlyList<ExcelChartSeries> series = snapshot.Data.Series;
            if (categories.Count == 0 || series.Count == 0) {
                return;
            }

            IReadOnlyList<double> xValues = GetScatterXValues(categories);
            (double minX, double maxX) = GetFiniteRange(xValues);
            (double minY, double maxY) = GetFiniteSeriesRange(series);
            for (int s = 0; s < series.Count; s++) {
                OfficeColor color = GetChartSeriesColor(s);
                var points = new List<OfficePoint>(categories.Count);
                for (int i = 0; i < categories.Count; i++) {
                    double yValue = GetSeriesValue(series[s], i);
                    double x = ToPlotX(xValues[i], minX, maxX, plotLeft, plotWidth);
                    double y = ToPlotY(yValue, minY, maxY, plotTop, plotHeight);
                    points.Add(new OfficePoint(x, y));
                }

                AddPointLine(drawing, points, color, 1.25D);
                for (int i = 0; i < points.Count; i++) {
                    OfficePoint point = points[i];
                    AddShape(drawing, OfficeShape.Ellipse(5D, 5D), point.X - 2.5D, point.Y - 2.5D, OfficeColor.White, color, 1.25D);
                }
            }
        }

        private static void AddRadarSeries(OfficeDrawing drawing, ExcelChartSnapshot snapshot, double width, double height) {
            IReadOnlyList<string> categories = snapshot.Data.Categories;
            IReadOnlyList<ExcelChartSeries> series = snapshot.Data.Series;
            if (categories.Count < 3 || series.Count == 0) {
                return;
            }

            double centerX = width / 2D;
            double centerY = height / 2D;
            double radius = Math.Max(36D, Math.Min(width - 52D, height - 42D) / 2D);
            var (min, max) = GetRadarValueRange(series);

            for (int ring = 1; ring <= 4; ring++) {
                double ringRadius = radius * ring / 4D;
                IReadOnlyList<OfficePoint> ringPoints = CreateRadarPoints(categories.Count, centerX, centerY, ringRadius);
                AddPolygonShape(drawing, ringPoints, null, OfficeColor.FromRgb(226, 232, 240), 0.5D);
            }

            IReadOnlyList<OfficePoint> outerPoints = CreateRadarPoints(categories.Count, centerX, centerY, radius);
            for (int i = 0; i < outerPoints.Count; i++) {
                OfficePoint point = outerPoints[i];
                double minX = Math.Min(centerX, point.X);
                double minY = Math.Min(centerY, point.Y);
                AddShape(
                    drawing,
                    OfficeShape.Line(centerX - minX, centerY - minY, point.X - minX, point.Y - minY),
                    minX,
                    minY,
                    null,
                    OfficeColor.FromRgb(203, 213, 225),
                    0.5D);
            }

            for (int s = 0; s < series.Count; s++) {
                OfficeColor color = GetChartSeriesColor(s);
                var points = new List<OfficePoint>(categories.Count);
                for (int i = 0; i < categories.Count; i++) {
                    double value = GetSeriesValue(series[s], i);
                    double pointRadius = radius * ToRadarRadiusRatio(value, min, max);
                    points.Add(CreateRadarPoint(i, categories.Count, centerX, centerY, pointRadius));
                }

                AddPolygonShape(drawing, points, color, color, 1D, 0.18D);
                for (int i = 0; i < points.Count; i++) {
                    OfficePoint point = points[i];
                    AddShape(drawing, OfficeShape.Ellipse(4D, 4D), point.X - 2D, point.Y - 2D, OfficeColor.White, color, 1D);
                }
            }
        }

        private static void AddPieSeries(OfficeDrawing drawing, ExcelChartSnapshot snapshot, double width, double height, bool doughnut) {
            IReadOnlyList<string> categories = snapshot.Data.Categories;
            IReadOnlyList<ExcelChartSeries> series = snapshot.Data.Series;
            if (categories.Count == 0 || series.Count == 0) {
                return;
            }

            ExcelChartSeries values = series[0];
            double total = 0D;
            for (int i = 0; i < categories.Count; i++) {
                double value = GetSeriesValue(values, i);
                if (!double.IsNaN(value) && !double.IsInfinity(value) && value > 0D) {
                    total += value;
                }
            }

            if (total <= 0D) {
                return;
            }

            double radius = Math.Max(36D, Math.Min(width - 48D, height - 36D) / 2D);
            double centerX = width / 2D;
            double centerY = height / 2D;
            double start = -Math.PI / 2D;
            for (int i = 0; i < categories.Count; i++) {
                double value = Math.Max(0D, GetSeriesValue(values, i));
                if (value <= 0D) {
                    continue;
                }

                double sweep = value / total * Math.PI * 2D;
                double end = start + sweep;
                var points = new List<OfficePoint> {
                    new OfficePoint(centerX, centerY)
                };
                int segments = Math.Max(2, (int)Math.Ceiling(sweep / (Math.PI / 18D)));
                for (int segment = 0; segment <= segments; segment++) {
                    double angle = start + (sweep * segment / segments);
                    points.Add(new OfficePoint(
                        centerX + Math.Cos(angle) * radius,
                        centerY + Math.Sin(angle) * radius));
                }

                AddPolygonShape(drawing, points, GetChartSeriesColor(i), OfficeColor.White, 0.5D);
                start = end;
            }

            if (doughnut) {
                double innerDiameter = radius * 1.02D;
                AddShape(
                    drawing,
                    OfficeShape.Ellipse(innerDiameter, innerDiameter),
                    centerX - innerDiameter / 2D,
                    centerY - innerDiameter / 2D,
                    OfficeColor.FromRgb(250, 252, 255),
                    null,
                    0D);
            }
        }

        private static string[][] CreateChartLegendRows(ExcelChartSnapshot snapshot) {
            if (IsPieChart(snapshot.ChartType) || IsDoughnutChart(snapshot.ChartType)) {
                return CreatePieChartLegendRows(snapshot);
            }

            var rows = new List<string[]> {
                new[] { "Series", "Values" }
            };

            foreach (ExcelChartSeries series in snapshot.Data.Series) {
                rows.Add(new[] {
                    series.Name,
                    string.Join(", ", series.Values.Select(value => value.ToString("0.##", CultureInfo.InvariantCulture)))
                });
            }

            return rows.ToArray();
        }

        private static string[][] CreatePieChartLegendRows(ExcelChartSnapshot snapshot) {
            var rows = new List<string[]> {
                new[] { "Category", "Value" }
            };

            IReadOnlyList<string> categories = snapshot.Data.Categories;
            IReadOnlyList<ExcelChartSeries> series = snapshot.Data.Series;
            ExcelChartSeries? values = series.Count > 0 ? series[0] : null;
            for (int i = 0; i < categories.Count; i++) {
                string category = string.IsNullOrWhiteSpace(categories[i])
                    ? "Slice " + (i + 1).ToString(CultureInfo.InvariantCulture)
                    : categories[i];
                rows.Add(new[] {
                    category,
                    values == null ? string.Empty : GetSeriesValue(values, i).ToString("0.##", CultureInfo.InvariantCulture)
                });
            }

            return rows.ToArray();
        }

        private static int GetChartLegendColorCount(ExcelChartSnapshot snapshot) {
            if (IsPieChart(snapshot.ChartType) || IsDoughnutChart(snapshot.ChartType)) {
                return snapshot.Data.Categories.Count;
            }

            return snapshot.Data.Series.Count;
        }

        private static PdfCore.PdfTableStyle CreateChartLegendStyle(int colorCount) {
            var style = new PdfCore.PdfTableStyle {
                HeaderRowCount = 1,
                FontSize = 8.5,
                HeaderFontSize = 8.5,
                CellPaddingX = 4,
                CellPaddingY = 2,
                BorderColor = PdfCore.PdfColor.FromRgb(203, 213, 225),
                HeaderFill = PdfCore.PdfColor.FromRgb(239, 246, 255),
                ColumnWidthWeights = new List<double> { 0.7D, 1.3D },
                AutoFitColumns = false,
                MaxWidth = 300D,
                SpacingAfter = 6
            };

            var fills = new Dictionary<(int Row, int Column), PdfCore.PdfColor>();
            for (int i = 0; i < colorCount; i++) {
                fills[(i + 1, 0)] = PdfCore.PdfColor.FromOfficeColor(GetChartSeriesColor(i));
            }
            style.CellFills = fills;
            return style;
        }

        private static double GetPositiveMax(IReadOnlyList<ExcelChartSeries> series) {
            double max = 0D;
            foreach (ExcelChartSeries item in series) {
                foreach (double value in item.Values) {
                    if (!double.IsNaN(value) && !double.IsInfinity(value) && value > max) {
                        max = value;
                    }
                }
            }

            return max <= 0D ? 1D : max;
        }

        private static double GetPositiveStackedMax(IReadOnlyList<ExcelChartSeries> series, int categoryCount) {
            double max = 0D;
            for (int i = 0; i < categoryCount; i++) {
                double total = GetPositiveCategoryTotal(series, i);
                if (total > max) {
                    max = total;
                }
            }

            return max <= 0D ? 1D : max;
        }

        private static double GetPositiveCategoryTotal(IReadOnlyList<ExcelChartSeries> series, int categoryIndex) {
            double total = 0D;
            for (int s = 0; s < series.Count; s++) {
                total += Math.Max(0D, GetSeriesValue(series[s], categoryIndex));
            }

            return total;
        }

        private static (double Min, double Max) GetStackedSeriesRange(IReadOnlyList<ExcelChartSeries> series, int categoryCount) {
            double min = 0D;
            double max = 0D;
            for (int category = 0; category < categoryCount; category++) {
                double positive = 0D;
                double negative = 0D;
                for (int s = 0; s < series.Count; s++) {
                    double value = GetSeriesValue(series[s], category);
                    if (value >= 0D) {
                        positive += value;
                    } else {
                        negative += value;
                    }
                }

                if (positive > max) max = positive;
                if (negative < min) min = negative;
            }

            return ExpandFlatRange(min, max);
        }

        private static double GetSeriesValue(ExcelChartSeries series, int index) {
            double value = index >= 0 && index < series.Values.Count ? series.Values[index] : 0D;
            return double.IsNaN(value) || double.IsInfinity(value) ? 0D : value;
        }

        private static double ToPlotY(double value, double max, double plotTop, double plotHeight) {
            double ratio = max <= 0D ? 0D : Math.Max(0D, value) / max;
            if (ratio > 1D) {
                ratio = 1D;
            }

            return plotTop + plotHeight - (plotHeight * ratio);
        }

        private static double ToPlotY(double value, double min, double max, double plotTop, double plotHeight) {
            double range = max - min;
            double ratio = range <= 0D ? 0.5D : (value - min) / range;
            if (ratio < 0D) {
                ratio = 0D;
            } else if (ratio > 1D) {
                ratio = 1D;
            }

            return plotTop + plotHeight - (plotHeight * ratio);
        }

        private static double ToPlotX(double value, double min, double max, double plotLeft, double plotWidth) {
            double range = max - min;
            double ratio = range <= 0D ? 0.5D : (value - min) / range;
            if (ratio < 0D) {
                ratio = 0D;
            } else if (ratio > 1D) {
                ratio = 1D;
            }

            return plotLeft + plotWidth * ratio;
        }

        private static IReadOnlyList<double> GetScatterXValues(IReadOnlyList<string> categories) {
            var values = new double[categories.Count];
            for (int i = 0; i < categories.Count; i++) {
                if (double.TryParse(categories[i], NumberStyles.Float, CultureInfo.InvariantCulture, out double value) &&
                    !double.IsNaN(value) &&
                    !double.IsInfinity(value)) {
                    values[i] = value;
                } else {
                    values[i] = i + 1D;
                }
            }

            return values;
        }

        private static IReadOnlyList<OfficePoint> CreateRadarPoints(int count, double centerX, double centerY, double radius) {
            var points = new List<OfficePoint>(count);
            for (int i = 0; i < count; i++) {
                points.Add(CreateRadarPoint(i, count, centerX, centerY, radius));
            }

            return points;
        }

        private static OfficePoint CreateRadarPoint(int index, int count, double centerX, double centerY, double radius) {
            double angle = -Math.PI / 2D + Math.PI * 2D * index / count;
            return new OfficePoint(centerX + Math.Cos(angle) * radius, centerY + Math.Sin(angle) * radius);
        }

        private static (double Min, double Max) GetFiniteSeriesRange(IReadOnlyList<ExcelChartSeries> series) {
            bool any = false;
            double min = 0D;
            double max = 0D;
            foreach (ExcelChartSeries item in series) {
                foreach (double value in item.Values) {
                    if (double.IsNaN(value) || double.IsInfinity(value)) {
                        continue;
                    }

                    if (!any) {
                        min = value;
                        max = value;
                        any = true;
                    } else {
                        if (value < min) min = value;
                        if (value > max) max = value;
                    }
                }
            }

            return any ? ExpandFlatRange(min, max) : (0D, 1D);
        }

        private static (double Min, double Max) GetFiniteRange(IReadOnlyList<double> values) {
            bool any = false;
            double min = 0D;
            double max = 0D;
            foreach (double value in values) {
                if (double.IsNaN(value) || double.IsInfinity(value)) {
                    continue;
                }

                if (!any) {
                    min = value;
                    max = value;
                    any = true;
                } else {
                    if (value < min) min = value;
                    if (value > max) max = value;
                }
            }

            return any ? ExpandFlatRange(min, max) : (0D, 1D);
        }

        private static (double Min, double Max) GetRadarValueRange(IReadOnlyList<ExcelChartSeries> series) {
            var (min, max) = GetFiniteSeriesRange(series);
            min = Math.Min(0D, min);
            max = Math.Max(0D, max);
            return ExpandFlatRange(min, max);
        }

        private static double ToRadarRadiusRatio(double value, double min, double max) {
            double range = max - min;
            double ratio = range <= 0D ? 0.5D : (value - min) / range;
            if (ratio < 0D) {
                return 0D;
            }

            if (ratio > 1D) {
                return 1D;
            }

            return ratio;
        }

        private static (double Min, double Max) ExpandFlatRange(double min, double max) {
            if (max > min) {
                return (min, max);
            }

            double padding = Math.Abs(min) > 1D ? Math.Abs(min) * 0.1D : 1D;
            return (min - padding, max + padding);
        }

        private static OfficeColor GetChartSeriesColor(int index) {
            switch (index % 6) {
                case 0:
                    return OfficeColor.FromRgb(31, 78, 121);
                case 1:
                    return OfficeColor.FromRgb(47, 111, 62);
                case 2:
                    return OfficeColor.FromRgb(184, 90, 35);
                case 3:
                    return OfficeColor.FromRgb(112, 48, 160);
                case 4:
                    return OfficeColor.FromRgb(37, 99, 235);
                default:
                    return OfficeColor.FromRgb(120, 113, 108);
            }
        }

        private static void AddShape(OfficeDrawing drawing, OfficeShape shape, double x, double y, OfficeColor? fill, OfficeColor? stroke, double strokeWidth) {
            shape.FillColor = fill;
            shape.StrokeColor = stroke;
            shape.StrokeWidth = strokeWidth;
            drawing.AddShape(shape, x, y);
        }

        private static void AddPolygonShape(OfficeDrawing drawing, IReadOnlyList<OfficePoint> points, OfficeColor? fill, OfficeColor? stroke, double strokeWidth, double? fillOpacity = null) {
            if (points.Count < 3) {
                return;
            }

            double minX = points[0].X;
            double minY = points[0].Y;
            double maxX = points[0].X;
            double maxY = points[0].Y;
            for (int i = 1; i < points.Count; i++) {
                OfficePoint point = points[i];
                if (point.X < minX) minX = point.X;
                if (point.Y < minY) minY = point.Y;
                if (point.X > maxX) maxX = point.X;
                if (point.Y > maxY) maxY = point.Y;
            }

            if (maxX <= minX || maxY <= minY) {
                return;
            }

            OfficeShape shape = OfficeShape.Polygon(points);
            shape.FillOpacity = fillOpacity;
            AddShape(drawing, shape, minX, minY, fill, stroke, strokeWidth);
        }

        private static void AddPointLine(OfficeDrawing drawing, IReadOnlyList<OfficePoint> points, OfficeColor color, double strokeWidth) {
            for (int i = 1; i < points.Count; i++) {
                OfficePoint previous = points[i - 1];
                OfficePoint current = points[i];
                if (previous.Equals(current)) {
                    continue;
                }

                double minX = Math.Min(previous.X, current.X);
                double minY = Math.Min(previous.Y, current.Y);
                AddShape(
                    drawing,
                    OfficeShape.Line(previous.X - minX, previous.Y - minY, current.X - minX, current.Y - minY),
                    minX,
                    minY,
                    null,
                    color,
                    strokeWidth);
            }
        }

        private static bool IsColumnChart(ExcelChartType type) {
            return type == ExcelChartType.ColumnClustered
                   || type == ExcelChartType.ColumnStacked
                   || type == ExcelChartType.ColumnStacked100
                   || type == ExcelChartType.Column3DClustered
                   || type == ExcelChartType.Column3DStacked
                   || type == ExcelChartType.Column3DStacked100;
        }

        private static bool IsBarChart(ExcelChartType type) {
            return type == ExcelChartType.BarClustered
                   || type == ExcelChartType.BarStacked
                   || type == ExcelChartType.BarStacked100
                   || type == ExcelChartType.Bar3DClustered
                   || type == ExcelChartType.Bar3DStacked
                   || type == ExcelChartType.Bar3DStacked100;
        }

        private static bool IsLineChart(ExcelChartType type) {
            return type == ExcelChartType.Line
                   || type == ExcelChartType.LineStacked
                   || type == ExcelChartType.LineStacked100
                   || type == ExcelChartType.Line3D;
        }

        private static bool IsAreaChart(ExcelChartType type) {
            return type == ExcelChartType.Area
                   || type == ExcelChartType.AreaStacked
                   || type == ExcelChartType.AreaStacked100
                   || type == ExcelChartType.Area3D
                   || type == ExcelChartType.Area3DStacked
                   || type == ExcelChartType.Area3DStacked100;
        }

        private static bool IsScatterChart(ExcelChartType type) {
            return type == ExcelChartType.Scatter;
        }

        private static bool IsRadarChart(ExcelChartType type) {
            return type == ExcelChartType.Radar;
        }

        private static bool IsStackedAreaChart(ExcelChartType type) {
            return type == ExcelChartType.AreaStacked
                   || type == ExcelChartType.Area3DStacked;
        }

        private static bool IsPercentStackedAreaChart(ExcelChartType type) {
            return type == ExcelChartType.AreaStacked100
                   || type == ExcelChartType.Area3DStacked100;
        }

        private static bool IsStackedBarOrColumnChart(ExcelChartType type) {
            return type == ExcelChartType.ColumnStacked
                   || type == ExcelChartType.Column3DStacked
                   || type == ExcelChartType.BarStacked
                   || type == ExcelChartType.Bar3DStacked;
        }

        private static bool IsPercentStackedBarOrColumnChart(ExcelChartType type) {
            return type == ExcelChartType.ColumnStacked100
                   || type == ExcelChartType.Column3DStacked100
                   || type == ExcelChartType.BarStacked100
                   || type == ExcelChartType.Bar3DStacked100;
        }

        private static bool IsPieChart(ExcelChartType type) {
            return type == ExcelChartType.Pie
                   || type == ExcelChartType.Pie3D
                   || type == ExcelChartType.PieOfPie
                   || type == ExcelChartType.BarOfPie;
        }

        private static bool IsDoughnutChart(ExcelChartType type) {
            return type == ExcelChartType.Doughnut;
        }

    }
}
