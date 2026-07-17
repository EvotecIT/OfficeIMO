using System.Globalization;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        /// <summary>
        /// Adds a histogram-compatible column chart after calculating bins from raw numeric observations.
        /// </summary>
        public ExcelChart AddHistogramChart(
            IEnumerable<double> values,
            int row,
            int column,
            string title = "Histogram",
            int? binCount = null,
            double? binWidth = null,
            int widthPixels = 640,
            int heightPixels = 360) {
            List<double> observations = MaterializeFiniteChartValues(values, nameof(values), allowNegative: true);
            if (binCount.HasValue && binWidth.HasValue) {
                throw new ArgumentException("Specify either binCount or binWidth, not both.");
            }
            if (binCount.HasValue && binCount.Value <= 0) {
                throw new ArgumentOutOfRangeException(nameof(binCount));
            }
            if (binWidth.HasValue && (!(binWidth.Value > 0) || double.IsInfinity(binWidth.Value))) {
                throw new ArgumentOutOfRangeException(nameof(binWidth));
            }

            double minimum = observations.Min();
            double maximum = observations.Max();
            double range = maximum - minimum;
            if (double.IsNaN(range) || double.IsInfinity(range)) {
                throw new ArgumentException("Histogram values span a range that is too large to aggregate safely.", nameof(values));
            }

            int effectiveBinCount;
            double effectiveBinWidth;
            if (minimum == maximum) {
                effectiveBinCount = 1;
                effectiveBinWidth = 1D;
            } else if (binWidth.HasValue) {
                effectiveBinWidth = binWidth.Value;
                double requestedBinCount = Math.Max(1D, Math.Ceiling(range / effectiveBinWidth));
                if (double.IsNaN(requestedBinCount) || double.IsInfinity(requestedBinCount) || requestedBinCount > 10000D) {
                    throw new ArgumentOutOfRangeException(nameof(binWidth), "A histogram cannot exceed 10,000 bins.");
                }

                effectiveBinCount = (int)requestedBinCount;
            } else {
                effectiveBinCount = binCount ?? Math.Max(1, (int)Math.Ceiling(Math.Sqrt(observations.Count)));
                effectiveBinWidth = range / effectiveBinCount;
            }
            if (effectiveBinCount > 10000) {
                throw new ArgumentOutOfRangeException(nameof(binCount), "A histogram cannot exceed 10,000 bins.");
            }
            if (!(effectiveBinWidth > 0) || double.IsInfinity(effectiveBinWidth)) {
                throw new ArgumentException("Histogram values and bin settings produce an unrepresentable bin width.", nameof(values));
            }

            var counts = new double[effectiveBinCount];
            foreach (double observation in observations) {
                int index = minimum == maximum || observation == maximum
                    ? effectiveBinCount - 1
                    : (int)Math.Floor((observation - minimum) / effectiveBinWidth);
                counts[Math.Max(0, Math.Min(effectiveBinCount - 1, index))]++;
            }

            var categories = new string[effectiveBinCount];
            for (int index = 0; index < effectiveBinCount; index++) {
                double lower = minimum + (index * effectiveBinWidth);
                double upper = minimum == maximum || index == effectiveBinCount - 1
                    ? maximum
                    : lower + effectiveBinWidth;
                if (double.IsInfinity(lower) || double.IsInfinity(upper)) {
                    throw new ArgumentException("Histogram values and bin settings produce an unrepresentable bin boundary.", nameof(values));
                }
                categories[index] = FormatChartRange(lower, upper);
            }

            var data = new ExcelChartData(categories, new[] { new ExcelChartSeries("Frequency", counts, seriesColorArgb: "4F46E5") });
            ExcelChart chart = AddChart(data, row, column, widthPixels, heightPixels, ExcelChartType.ColumnClustered, title);
            return chart.HideLegend()
                .SetCategoryAxisTitle("Bin")
                .SetValueAxisTitle("Frequency")
                .SetValueAxisNumberFormat("0", sourceLinked: false)
                .SetValueAxisGridlines(showMajor: true, showMinor: false, lineColor: "E5E7EB", lineWidthPoints: 0.5);
        }

        /// <summary>
        /// Adds a Pareto-compatible column and cumulative-percentage line chart.
        /// </summary>
        public ExcelChart AddParetoChart(
            IEnumerable<string> categories,
            IEnumerable<double> values,
            int row,
            int column,
            string title = "Pareto",
            int widthPixels = 720,
            int heightPixels = 360) {
            List<(string Category, double Value)> points = MaterializeCategoryChartValues(categories, values, nameof(values), allowNegative: false)
                .OrderByDescending(point => point.Value)
                .ThenBy(point => point.Category, StringComparer.OrdinalIgnoreCase)
                .ToList();
            double total = points.Sum(point => point.Value);
            if (!(total > 0) || double.IsInfinity(total)) {
                throw new ArgumentException("Pareto values must contain a finite positive total.", nameof(values));
            }

            double running = 0;
            var cumulative = new double[points.Count];
            for (int index = 0; index < points.Count; index++) {
                running += points[index].Value;
                if (double.IsNaN(running) || double.IsInfinity(running)) {
                    throw new ArgumentException("Pareto values are too large to aggregate safely.", nameof(values));
                }

                cumulative[index] = running / total;
            }

            var data = new ExcelChartData(
                points.Select(point => point.Category),
                new[] {
                    new ExcelChartSeries("Value", points.Select(point => point.Value), ExcelChartType.ColumnClustered, ExcelChartAxisGroup.Primary, "4F46E5"),
                    new ExcelChartSeries("Cumulative %", cumulative, ExcelChartType.Line, ExcelChartAxisGroup.Secondary, "F59E0B")
                });
            ExcelChart chart = AddChart(data, row, column, widthPixels, heightPixels, ExcelChartType.ColumnClustered, title);
            return chart.SetLegend(C.LegendPositionValues.Bottom)
                .SetSeriesLineColor(1, "F59E0B", 2)
                .SetSeriesMarker(1, C.MarkerStyleValues.Circle, size: 6, fillColor: "F59E0B", lineColor: "F59E0B")
                .SetValueAxisNumberFormat("0%", sourceLinked: false, ExcelChartAxisGroup.Secondary)
                .SetValueAxisScale(minimum: 0, maximum: 1, majorUnit: 0.2, axisGroup: ExcelChartAxisGroup.Secondary)
                .SetValueAxisGridlines(showMajor: true, showMinor: false, lineColor: "E5E7EB", lineWidthPoints: 0.5);
        }

        /// <summary>
        /// Adds a centered funnel-compatible stacked bar chart from category values.
        /// </summary>
        public ExcelChart AddFunnelChart(
            IEnumerable<string> categories,
            IEnumerable<double> values,
            int row,
            int column,
            string title = "Funnel",
            bool sortDescending = true,
            int widthPixels = 640,
            int heightPixels = 360) {
            IEnumerable<(string Category, double Value)> source = MaterializeCategoryChartValues(categories, values, nameof(values), allowNegative: false);
            List<(string Category, double Value)> points = (sortDescending
                ? source.OrderByDescending(point => point.Value).ThenBy(point => point.Category, StringComparer.OrdinalIgnoreCase)
                : source).ToList();
            double maximum = points.Max(point => point.Value);
            if (!(maximum > 0)) {
                throw new ArgumentException("Funnel values must contain a positive value.", nameof(values));
            }

            double[] spacer = points.Select(point => (maximum - point.Value) / 2D).ToArray();
            double[] actual = points.Select(point => point.Value).ToArray();
            var data = new ExcelChartData(
                points.Select(point => point.Category),
                new[] {
                    new ExcelChartSeries("Left spacer", spacer),
                    new ExcelChartSeries("Value", actual, seriesColorArgb: "0EA5E9"),
                    new ExcelChartSeries("Right spacer", spacer)
                });
            ExcelChart chart = AddChart(data, row, column, widthPixels, heightPixels, ExcelChartType.BarStacked, title);
            chart.ApplyAuthoredSeriesStyles(data.Series, new[] { false, true, false });
            return chart.SetSeriesNoFill(0)
                .SetSeriesNoFill(2)
                .SetSeriesDataLabels(1, showValue: true, position: C.DataLabelPositionValues.Center, numberFormat: "#,##0")
                .SetCategoryAxisReverseOrder()
                .SetValueAxisGridlines(showMajor: false, showMinor: false);
        }

        /// <summary>
        /// Adds a waterfall-compatible stacked column chart from signed changes and an optional final total.
        /// </summary>
        public ExcelChart AddWaterfallChart(
            IEnumerable<string> categories,
            IEnumerable<double> changes,
            int row,
            int column,
            string title = "Waterfall",
            bool includeTotal = true,
            string totalLabel = "Total",
            int widthPixels = 720,
            int heightPixels = 360) {
            List<(string Category, double Value)> points = MaterializeCategoryChartValues(categories, changes, nameof(changes), allowNegative: true);
            int count = points.Count + (includeTotal ? 1 : 0);
            var chartCategories = new List<string>(count);
            var baseline = new double[count];
            var increases = new double[count];
            var decreases = new double[count];
            var totals = new double[count];
            double running = 0;
            for (int index = 0; index < points.Count; index++) {
                (string category, double change) = points[index];
                double next = running + change;
                if (double.IsNaN(next) || double.IsInfinity(next)) {
                    throw new ArgumentException("Waterfall changes are too large to aggregate safely.", nameof(changes));
                }

                double roundingTolerance = 3.552713678800501E-15D
                    * Math.Max(1D, Math.Max(Math.Abs(running), Math.Abs(change)));
                if (next < 0 && next >= -roundingTolerance) {
                    next = 0;
                }

                if (next < 0) {
                    throw new ArgumentException("The compatible waterfall recipe requires a non-negative running total.", nameof(changes));
                }

                chartCategories.Add(category);
                if (change >= 0) {
                    baseline[index] = running;
                    increases[index] = change;
                } else {
                    baseline[index] = next;
                    decreases[index] = -change;
                }
                running = next;
            }

            if (includeTotal) {
                chartCategories.Add(string.IsNullOrWhiteSpace(totalLabel) ? "Total" : totalLabel.Trim());
                totals[count - 1] = running;
            }

            var series = new List<ExcelChartSeries> {
                new ExcelChartSeries("Baseline", baseline),
                new ExcelChartSeries("Increase", increases, seriesColorArgb: "16A34A"),
                new ExcelChartSeries("Decrease", decreases, seriesColorArgb: "DC2626")
            };
            var visibleSeriesStyles = new List<bool> { false, true, true };
            if (includeTotal) {
                series.Add(new ExcelChartSeries("Total", totals, seriesColorArgb: "2563EB"));
                visibleSeriesStyles.Add(true);
            }

            var data = new ExcelChartData(chartCategories, series);
            ExcelChart chart = AddChart(data, row, column, widthPixels, heightPixels, ExcelChartType.ColumnStacked, title);
            chart.ApplyAuthoredSeriesStyles(data.Series, visibleSeriesStyles);
            chart.SetSeriesNoFill(0)
                .SetValueAxisGridlines(showMajor: true, showMinor: false, lineColor: "E5E7EB", lineWidthPoints: 0.5)
                .SetValueAxisNumberFormat("#,##0", sourceLinked: false);

            for (int index = 0; index < points.Count; index++) {
                if (points[index].Value > 0) {
                    chart.SetSeriesDataLabelForPoint(1, index, showValue: true,
                        position: C.DataLabelPositionValues.OutsideEnd, numberFormat: "+#,##0");
                } else if (points[index].Value < 0) {
                    chart.SetSeriesDataLabelForPoint(2, index, showValue: true,
                        position: C.DataLabelPositionValues.OutsideEnd, numberFormat: "-#,##0");
                }
            }

            if (includeTotal) {
                chart.SetSeriesDataLabelForPoint(3, count - 1, showValue: true,
                    position: C.DataLabelPositionValues.OutsideEnd, numberFormat: "#,##0");
            }

            return chart;
        }

        private static List<double> MaterializeFiniteChartValues(IEnumerable<double> values, string parameterName, bool allowNegative) {
            if (values == null) {
                throw new ArgumentNullException(parameterName);
            }

            List<double> materialized = values.ToList();
            if (materialized.Count == 0) {
                throw new ArgumentException("At least one value is required.", parameterName);
            }

            if (materialized.Any(value => double.IsNaN(value) || double.IsInfinity(value) || (!allowNegative && value < 0))) {
                throw new ArgumentException(allowNegative
                    ? "Chart values must be finite numbers."
                    : "Chart values must be finite non-negative numbers.", parameterName);
            }

            return materialized;
        }

        private static List<(string Category, double Value)> MaterializeCategoryChartValues(
            IEnumerable<string> categories,
            IEnumerable<double> values,
            string parameterName,
            bool allowNegative) {
            if (categories == null) {
                throw new ArgumentNullException(nameof(categories));
            }

            List<string> materializedCategories = categories.Select(category => category ?? string.Empty).ToList();
            List<double> materializedValues = MaterializeFiniteChartValues(values, parameterName, allowNegative);
            if (materializedCategories.Count != materializedValues.Count) {
                throw new ArgumentException("Categories and values must contain the same number of items.", parameterName);
            }

            return materializedCategories
                .Select((category, index) => (category, materializedValues[index]))
                .ToList();
        }

        private static string FormatChartRange(double lower, double upper) {
            string lowerText = lower.ToString("G15", CultureInfo.InvariantCulture);
            string upperText = upper.ToString("G15", CultureInfo.InvariantCulture);
            if (lower != upper && string.Equals(lowerText, upperText, StringComparison.Ordinal)) {
                lowerText = lower.ToString("G17", CultureInfo.InvariantCulture);
                upperText = upper.ToString("G17", CultureInfo.InvariantCulture);
            }

            return lowerText + " – " + upperText;
        }
    }
}
