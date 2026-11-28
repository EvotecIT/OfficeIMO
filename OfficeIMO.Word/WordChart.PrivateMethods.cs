using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.OpenXml.Internal;

namespace OfficeIMO.Word {
    // Holds private helper methods for chart construction.
    public partial class WordChart {
        /// <summary>The pie, 3-D pie or doughnut group of the chart, or none.</summary>
        private OfficeXmlNode PieFamily {
            get {
                OfficeXmlNode plot = PlotAreaOf(_chart);
                if (plot.IsNone) return default;
                OfficeXmlNode family = plot.Child(NsC, "pieChart");
                if (family.IsNone) family = plot.Child(NsC, "pie3DChart");
                if (family.IsNone) family = plot.Child(NsC, "doughnutChart");
                return family;
            }
        }

        private static OfficeXmlChildOrder PieFamilyOrder(OfficeXmlNode family) =>
            family.LocalName == "doughnutChart" ? CDoughnutChartView.Order : CPieChartView.Order;

        /// <summary>The pie series' <c>c:cat</c>, or a detached one with an empty string literal when it has none.</summary>
        private OfficeXmlNode InitializeCategoryAxisData(OfficeXmlNode pieChartSeries) {
            OfficeXmlNode categoryAxis = pieChartSeries.IsNone ? default : pieChartSeries.Child(NsC, "cat");
            // If CategoryAxisData does not exist, create it
            if (categoryAxis.IsNone) {
                categoryAxis = pieChartSeries.IsNone ? NewC(PlotSeriesContext, "cat") : NewC(pieChartSeries, "cat");
                AddC(categoryAxis, "strLit");
            }
            return categoryAxis;
        }

        /// <summary>The element new detached series data is made for when the chart has no pie series to hold it.</summary>
        private OfficeXmlNode PlotSeriesContext => ChartSpaceNode;

        private OfficeXmlNode InitializeNumberLiteral(OfficeXmlNode values) {
            OfficeXmlNode valueAxis = PlotChild("valAx");
            OfficeXmlNode axisValues = valueAxis.IsNone ? default : valueAxis.Child(NsC, "val");
            OfficeXmlNode literal = axisValues.IsNone ? default : axisValues.Child(NsC, "numLit");
            // If NumberLiteral does not exist, create it
            if (literal.IsNone) {
                literal = NewC(values, "numLit");
                AddText(literal, NsC, "formatCode", "General");
            }
            return literal;
        }

        private OfficeXmlNode InitializeValues(OfficeXmlNode pieChartSeries) {
            OfficeXmlNode values = pieChartSeries.IsNone ? default : pieChartSeries.Child(NsC, "val");
            if (!values.IsNone) return values;
            values = pieChartSeries.IsNone ? NewC(PlotSeriesContext, "val") : NewC(pieChartSeries, "val");
            values.Append(InitializeNumberLiteral(values));
            return values;
        }

        /// <summary>The pie family's first series, created when the family has none, or none without a pie family.</summary>
        private OfficeXmlNode InitializePieChartSeries() {
            OfficeXmlNode chartElement = PieFamily;
            if (chartElement.IsNone) return default;
            OfficeXmlNode pieChartSeries = chartElement.Child(NsC, "ser");
            if (pieChartSeries.IsNone) {
                pieChartSeries = OfficeChartXml.SetOrdered(chartElement, CreatePieChartSeries(chartElement, _index, "Title?"),
                    PieFamilyOrder(chartElement));
            }
            return pieChartSeries;
        }

        private static OfficeXmlNode CreatePieChartSeries(OfficeXmlNode family, uint index, string series) {
            OfficeXmlNode pieChartSeries = NewC(family, "ser");
            AddC(pieChartSeries, "idx", index);
            AddC(pieChartSeries, "order", index);
            AddSeriesText(pieChartSeries, series);
            return pieChartSeries;
        }

        /// <summary>
        /// Adds single category to charts
        /// </summary>
        /// <param name="category">The category.</param>
        private void AddSingleCategory(string category) {
            OfficeXmlNode pieChartSeries = InitializePieChartSeries();

            OfficeXmlNode categoryAxis = InitializeCategoryAxisData(pieChartSeries);

            OfficeXmlNode numericCategories = categoryAxis.Child(NsC, "numLit");
            if (!numericCategories.IsNone) {
                _lastAppendedCategoryPoint = AppendLiteralSlicePoint(numericCategories, NewPoint(numericCategories, _currentIndexCategory, category));
                OfficeXmlNode count = numericCategories.Child(NsC, "ptCount");
                if (!count.IsNone) count.SetAttribute("val", OfficeChartXml.Text(_currentIndexCategory + 1));
                else OfficeChartXml.SetOrdered(numericCategories, NewC(numericCategories, "ptCount", _currentIndexCategory + 1), CNumDataView.Order);
                _currentIndexCategory++;
                return;
            }
            OfficeXmlNode stringLiteral = categoryAxis.Child(NsC, "strLit");
            // If StringLiteral does not exist, create it
            if (stringLiteral.IsNone) {
                stringLiteral = AddC(categoryAxis, "strLit");
            }
            _lastAppendedCategoryPoint = AppendLiteralSlicePoint(stringLiteral, NewPoint(stringLiteral, _currentIndexCategory, category));
            // Update the PointCount
            OfficeXmlNode pointCount = stringLiteral.Child(NsC, "ptCount");
            if (!pointCount.IsNone) {
                pointCount.SetAttribute("val", OfficeChartXml.Text(_currentIndexCategory + 1));
            } else {
                stringLiteral.InsertBefore(NewC(stringLiteral, "ptCount", _currentIndexCategory + 1), stringLiteral.FirstChild);
            }
            // Increment the current index
            _currentIndexCategory++;

            if (!pieChartSeries.IsNone && pieChartSeries.Child(NsC, "cat").IsNone) {
                categoryAxis = OfficeChartXml.SetOrdered(pieChartSeries, categoryAxis, CPieSerView.Order);
                RefreshAppendedCategoryPoint(categoryAxis);
            }
        }

        /// <summary>
        /// Points the last appended category at its element in <paramref name="categoryAxis"/>, as inserting a detached
        /// category list can hold it as new elements.
        /// </summary>
        private void RefreshAppendedCategoryPoint(OfficeXmlNode categoryAxis) {
            OfficeXmlNode literal = categoryAxis.Child(NsC, "numLit");
            if (literal.IsNone) literal = categoryAxis.Child(NsC, "strLit");
            OfficeXmlNode last = literal.IsNone ? default : literal.Children(NsC, "pt").Last();
            _lastAppendedCategoryPoint = last;
        }

        /// <summary>
        /// Adds the single value to charts
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="data">The data.</param>
        private void AddSingleValue<T>(T data) {
            // Initialize the PieChartSeries
            OfficeXmlNode pieChartSeries = InitializePieChartSeries();
            // Initialize the Values
            OfficeXmlNode values = InitializeValues(pieChartSeries);
            // Initialize the NumberLiteral
            OfficeXmlNode literal = values.Child(NsC, "numLit");
            if (literal.IsNone) literal = InitializeNumberLiteral(values);

            // Ensure decimal values use invariant culture (period as decimal separator) for OpenXML compatibility
            string valueText = data is double d ? OfficeIMO.Core.Internal.OfficeRoundTripText.Format(d, System.Globalization.CultureInfo.InvariantCulture) :
                              data is float f ? OfficeIMO.Core.Internal.OfficeRoundTripText.Format(f, System.Globalization.CultureInfo.InvariantCulture) :
                              data?.ToString() ?? "0";

            _lastAppendedValuePoint = AppendLiteralSlicePoint(literal, NewPoint(literal, _currentIndexValues, valueText));
            // Update the PointCount
            OfficeXmlNode pointCount = literal.Child(NsC, "ptCount");
            if (!pointCount.IsNone) {
                pointCount.SetAttribute("val", OfficeChartXml.Text(_currentIndexValues + 1));
            } else {
                OfficeXmlNode formatCode = literal.Child(NsC, "formatCode");
                OfficeXmlNode count = NewC(literal, "ptCount", _currentIndexValues + 1);
                if (formatCode.IsNone) literal.InsertBefore(count, literal.FirstChild);
                else formatCode.InsertAfterSelf(count);
            }
            // Increment the current index
            _currentIndexValues++;
            // add values to the series if it does not exist
            if (!pieChartSeries.IsNone && pieChartSeries.Child(NsC, "val").IsNone) {
                values = OfficeChartXml.SetOrdered(pieChartSeries, values, CPieSerView.Order);
                OfficeXmlNode heldLiteral = values.Child(NsC, "numLit");
                _lastAppendedValuePoint = heldLiteral.IsNone ? default : heldLiteral.Children(NsC, "pt").Last();
            }
        }

        private static void CreatePieChart(OfficeXmlNode chart) {
            OfficeXmlNode plotArea = PlotAreaOf(chart);
            OfficeXmlNode pieChart = NewC(plotArea, "pieChart");
            AddC(pieChart, "varyColors", true);
            pieChart.Append(AddDataLabel(pieChart));
            pieChart.DeclareNamespace("c", ChartNamespaceUri);
            plotArea.Append(pieChart);
        }

        /// <summary>
        /// Chart groups precede the axes in c:plotArea, so a group added to an existing chart goes before them
        /// and shares the axes already there.
        /// </summary>
        private static void AddChartGroup(OfficeXmlNode plotArea, OfficeXmlNode group, OfficeXmlNode categoryAxis, OfficeXmlNode valueAxis) {
            OfficeXmlNode firstAxis = FirstAxis(plotArea);
            if (firstAxis.IsNone) {
                plotArea.Append(group);
                plotArea.Append(categoryAxis);
                plotArea.Append(valueAxis);
                return;
            }

            OfficeXmlNode sharedCategory = plotArea.Child(NsC, "catAx");
            OfficeXmlNode sharedValue = plotArea.Child(NsC, "valAx");
            group = plotArea.InsertBefore(group, firstAxis);
            OfficeXmlNode sharedCategoryId = sharedCategory.IsNone ? default : sharedCategory.Child(NsC, "axId");
            OfficeXmlNode sharedValueId = sharedValue.IsNone ? default : sharedValue.Child(NsC, "axId");
            if (sharedCategoryId.Val() != null && sharedValueId.Val() != null) {
                List<OfficeXmlNode> ids = group.Children(NsC, "axId").ToList();
                ids[0].SetAttribute("val", OfficeChartXml.Text(sharedCategoryId.UIntVal()!.Value));
                ids[1].SetAttribute("val", OfficeChartXml.Text(sharedValueId.UIntVal()!.Value));
                return;
            }

            OfficeXmlNode lastAxis = default;
            foreach (OfficeXmlNode child in plotArea.Children()) {
                if (IsAxis(child)) lastAxis = child;
            }
            lastAxis.InsertAfterSelf(valueAxis);
            lastAxis.InsertAfterSelf(categoryAxis);
        }

        private static OfficeXmlNode FirstAxis(OfficeXmlNode plotArea) {
            foreach (OfficeXmlNode child in plotArea.Children()) {
                if (IsAxis(child)) return child;
            }
            return default;
        }

        private void GenerateChartBar(OfficeXmlNode chart) {
            uint catId = GenerateAxisId();
            uint valId = GenerateAxisId();

            OfficeXmlNode plotArea = PlotAreaOf(chart);
            OfficeXmlNode barChart = CreateBarChart(plotArea, catId, valId);
            OfficeXmlNode categoryAxis = AddCategoryAxisInternal(plotArea, catId, valId, "l");
            OfficeXmlNode valueAxis = AddValueAxisInternal(plotArea, valId, catId, "b");
            AddChartGroup(plotArea, barChart, categoryAxis, valueAxis);
        }

        private static OfficeXmlNode CreateBarChart(OfficeXmlNode plotArea, uint catAxisId, uint valAxisId, string barDirection = "bar") {
            OfficeXmlNode barChart = NewC(plotArea, "barChart");
            barChart.DeclareNamespace("c", ChartNamespaceUri);

            // Append elements in correct OpenXML schema order:
            // 1. barDir [1..1] - Bar Direction
            // 2. grouping [0..1] - Bar Grouping
            // 3. varyColors [0..1] - (not used)
            // 4. ser [0..*] - Bar Chart Series (added later when series are created)
            // 5. dLbls [0..1] - Data Labels
            // 6. gapWidth [0..1] - Gap Width
            // 7. overlap [0..1] - Overlap
            // 8. serLines [0..*] - Series Lines (not used)
            // 9. axId [2..2] - Axis ID
            // 10. extLst [0..1] - Extension List (not used)
            AddC(barChart, "barDir", barDirection);
            AddC(barChart, "grouping", "clustered");
            barChart.Append(AddDataLabel(barChart));
            AddC(barChart, "gapWidth", OfficeChartXml.Text(OfficeTheme.Default.ChartStyle.BarGapPercent));
            AddC(barChart, "overlap", "0");
            AddC(barChart, "axId", catAxisId).DeclareNamespace("c", ChartNamespaceUri);
            AddC(barChart, "axId", valAxisId).DeclareNamespace("c", ChartNamespaceUri);
            return barChart;
        }

        /// <summary>Builds a new chart with one group, attaches it and applies the chart-level defaults.</summary>
        private void CreateNewChart(Action<OfficeXmlNode> generate, bool axisTitles = true) {
            OfficeXmlNode chart = GenerateChart();
            generate(chart);
            AttachNewChart(chart);
            EnsureFramelessChartArea();

            // since the title may have changed, we need to update it
            UpdateTitle();
            if (axisTitles) UpdateAxisTitles();
        }

        private void EnsureChartExistsPie() {
            // minimum required to create chart
            if (_chart.IsNone) CreateNewChart(CreatePieChart);
        }

        private void EnsureChartExistsBar() {
            if (_chart.IsNone) {
                CreateNewChart(GenerateChartBar);
            } else {
                AttachCategoryGroup("barChart", (plot, catAxisId, valAxisId) => CreateBarChart(plot, catAxisId, valAxisId), GenerateChartBar);
            }
        }

        private void EnsureChartExistsArea() {
            if (_chart.IsNone) {
                CreateNewChart(GenerateAreaChart);
            } else {
                AttachCategoryGroup("areaChart", CreateAreaChart, GenerateAreaChart);
            }
        }

        /// <summary>
        /// Adds a bar, line or area group to an existing chart on the category and value axes of the
        /// chart's existing bar, line or area group, so loaded charts accept the same combinations as new ones.
        /// </summary>
        private void AttachCategoryGroup(string groupName, Func<OfficeXmlNode, uint, uint, OfficeXmlNode> create, Action<OfficeXmlNode> generate) {
            OfficeXmlNode plot = PlotAreaOf(_chart);
            if (plot.IsNone) plot = OfficeChartXml.SetOrdered(_chart, NewC(_chart, "plotArea"), CChartView.Order);
            if (!plot.Child(NsC, groupName).IsNone) return;
            OfficeXmlNode shared = plot.Child(NsC, "barChart");
            if (shared.IsNone) shared = plot.Child(NsC, "lineChart");
            if (shared.IsNone) shared = plot.Child(NsC, "areaChart");
            List<OfficeXmlNode>? ids = shared.IsNone ? null : shared.Children(NsC, "axId").ToList();
            if (ids is { Count: >= 2 } && ids[0].Val() != null && ids[1].Val() != null) {
                OfficeXmlNode group = create(plot, ids[0].UIntVal()!.Value, ids[1].UIntVal()!.Value);
                plot.InsertBefore(group, FirstAxis(plot));
                return;
            }
            RejectOtherChartGroup(groupName, groupName.Replace("Chart", string.Empty).ToLowerInvariant());
            generate(_chart);
        }

        /// <summary>Rejects a series kind that cannot share the plot area with the chart's existing group.</summary>
        private void RejectOtherChartGroup(string groupName, string kind) {
            OfficeXmlNode plot = PlotAreaOf(_chart);
            if (plot.IsNone || !plot.Child(NsC, groupName).IsNone) return;
            OfficeXmlNode existing = default;
            foreach (OfficeXmlNode child in plot.Children()) {
                if (IsChartGroup(child)) {
                    existing = child;
                    break;
                }
            }
            if (existing.IsNone) return;
            throw new NotSupportedException("A " + kind + " series cannot be added to a chart that already has a " +
                existing.LocalName.Replace("Chart", string.Empty) + " chart group. Bar, line and area series can be combined on shared axes.");
        }

        private void EnsureChartExistsLine() {
            if (_chart.IsNone) {
                CreateNewChart(GenerateLineChart);
            } else {
                AttachCategoryGroup("lineChart", CreateLineChart, GenerateLineChart);
            }
        }

        /// <summary>Starts a series with its index, order and literal name.</summary>
        private static OfficeXmlNode NewSeries(OfficeXmlNode group, uint index, string name) {
            OfficeXmlNode series = NewC(group, "ser");
            AddC(series, "idx", index);
            AddC(series, "order", index);
            AddSeriesText(series, name);
            return series;
        }

        private OfficeXmlNode AddBarChartSeries<T>(OfficeXmlNode group, uint index, string series, OfficeIMO.OfficeColor color, List<string> categories, List<T> data) {
            OfficeXmlNode barChartSeries = NewSeries(group, index, series);
            barChartSeries.Append(AddShapeProperties(barChartSeries, color));
            AddC(barChartSeries, "invertIfNegative");
            barChartSeries.Append(AddCategoryAxisData(barChartSeries, categories));
            barChartSeries.Append(AddValuesAxisData(barChartSeries, data));
            return barChartSeries;
        }

        /// <summary>A detached <c>c:spPr</c> with a solid fill, made for <paramref name="context"/>.</summary>
        private OfficeXmlNode AddShapeProperties(OfficeXmlNode context, OfficeIMO.OfficeColor color) {
            OfficeXmlNode chartShapeProperties = NewC(context, "spPr");
            chartShapeProperties.Append(CreateChartFill(chartShapeProperties, color));
            return chartShapeProperties;
        }

        private static OfficeXmlNode CreateLineChart(OfficeXmlNode plotArea, uint catAxisId, uint valAxisId) {
            OfficeXmlNode lineChart = NewC(plotArea, "lineChart");
            lineChart.DeclareNamespace("c", ChartNamespaceUri);
            AddC(lineChart, "grouping", "standard");
            lineChart.Append(AddDataLabel(lineChart));
            AddC(lineChart, "axId", catAxisId).DeclareNamespace("c", ChartNamespaceUri);
            AddC(lineChart, "axId", valAxisId).DeclareNamespace("c", ChartNamespaceUri);
            return lineChart;
        }

        private void GenerateLineChart(OfficeXmlNode chart) {
            uint catId = GenerateAxisId();
            uint valId = GenerateAxisId();

            OfficeXmlNode plotArea = PlotAreaOf(chart);
            OfficeXmlNode lineChart = CreateLineChart(plotArea, catId, valId);
            OfficeXmlNode categoryAxis = AddCategoryAxisInternal(plotArea, catId, valId, "b");
            OfficeXmlNode valueAxis = AddValueAxisInternal(plotArea, valId, catId, "l");
            AddChartGroup(plotArea, lineChart, categoryAxis, valueAxis);
        }

        private OfficeXmlNode AddLineChartSeries<T>(OfficeXmlNode group, uint index, string series, OfficeIMO.OfficeColor color, List<string> categories, List<T> data) {
            // Append elements in correct OpenXML schema order for LineChartSeries:
            // 1. idx, 2. order, 3. tx (seriesText), 4. spPr (chartShapeProperties),
            // 5. marker, 6. pictureOptions, 7. dPt, 8. dLbls, 9. trendline, 10. errBars,
            // 11. cat (categoryAxisData), 12. val (values), 13. smooth, 14. extLst
            // InvertIfNegative is not valid for LineChartSeries according to OpenXML schema.
            OfficeXmlNode lineChartSeries = NewSeries(group, index, series);
            lineChartSeries.Append(AddShapeProperties(lineChartSeries, color));
            AddC(AddC(lineChartSeries, "marker"), "symbol", "none");
            lineChartSeries.Append(AddCategoryAxisData(lineChartSeries, categories));
            lineChartSeries.Append(AddValuesAxisData(lineChartSeries, data));
            return lineChartSeries;
        }

        private static OfficeXmlNode CreateAreaChart(OfficeXmlNode plotArea, uint catAxisId, uint valAxisId) {
            OfficeXmlNode chart = NewC(plotArea, "areaChart");
            chart.DeclareNamespace("c", ChartNamespaceUri);
            AddC(chart, "grouping", "standard");
            chart.Append(AddDataLabel(chart));
            AddC(chart, "axId", catAxisId).DeclareNamespace("c", ChartNamespaceUri);
            AddC(chart, "axId", valAxisId).DeclareNamespace("c", ChartNamespaceUri);
            return chart;
        }

        private void GenerateAreaChart(OfficeXmlNode chart) {
            uint catId = GenerateAxisId();
            uint valId = GenerateAxisId();

            OfficeXmlNode plotArea = PlotAreaOf(chart);
            OfficeXmlNode areaChart = CreateAreaChart(plotArea, catId, valId);
            OfficeXmlNode categoryAxis = AddCategoryAxisInternal(plotArea, catId, valId, "b");
            OfficeXmlNode valueAxis = AddValueAxisInternal(plotArea, valId, catId, "l");
            AddChartGroup(plotArea, areaChart, categoryAxis, valueAxis);
        }

        private OfficeXmlNode AddAreaChartSeries<T>(OfficeXmlNode group, uint index, string series, OfficeIMO.OfficeColor color, List<string> categories, List<T> data) {
            OfficeXmlNode areaChartSeries = NewSeries(group, index, series);
            areaChartSeries.Append(AddShapeProperties(areaChartSeries, color));
            areaChartSeries.Append(AddCategoryAxisData(areaChartSeries, categories));
            areaChartSeries.Append(AddValuesAxisData(areaChartSeries, data));
            return areaChartSeries;
        }

        private static uint GenerateAxisId() => (uint)System.Threading.Interlocked.Increment(ref _axisIdSeed);

        /// <summary>
        /// Inserts a series into its chart group after the grouping or vary-colors setting, or before the data labels,
        /// gap depth or axis ids, and returns it as it is now held.
        /// </summary>
        private static OfficeXmlNode InsertSeries(OfficeXmlNode chartElement, OfficeXmlNode series) {
            // For Line3DChart and other charts, series should come after grouping/varyColors but before dLbls/gapDepth/axId

            // Try to insert after VaryColors first
            OfficeXmlNode varyColors = chartElement.Child(NsC, "varyColors");
            if (!varyColors.IsNone) return varyColors.InsertAfterSelf(series);

            // If no VaryColors, try to insert after Grouping. A bar group's grouping is a bar grouping, which this
            // doesn't look for, so bar series go before the data labels and keep the order they were added in.
            OfficeXmlNode grouping = chartElement.LocalName is "barChart" or "bar3DChart" ? default : chartElement.Child(NsC, "grouping");
            if (!grouping.IsNone) return grouping.InsertAfterSelf(series);

            // If no grouping, try to insert before DataLabels
            OfficeXmlNode dataLabels = chartElement.Child(NsC, "dLbls");
            if (!dataLabels.IsNone) return chartElement.InsertBefore(series, dataLabels);

            // If no DataLabels, try to insert before GapDepth
            OfficeXmlNode gapDepth = chartElement.Child(NsC, "gapDepth");
            if (!gapDepth.IsNone) return chartElement.InsertBefore(series, gapDepth);

            // As a last resort, insert before the first AxisId; if there is none, append it
            return chartElement.InsertBefore(series, chartElement.Child(NsC, "axId"));
        }

        private static OfficeXmlNode CreateScatterChart(OfficeXmlNode plotArea, uint xAxisId, uint yAxisId) {
            OfficeXmlNode chart = NewC(plotArea, "scatterChart");
            chart.DeclareNamespace("c", ChartNamespaceUri);
            AddC(chart, "scatterStyle", "marker");
            chart.Append(AddDataLabel(chart));
            AddC(chart, "axId", xAxisId).DeclareNamespace("c", ChartNamespaceUri);
            AddC(chart, "axId", yAxisId).DeclareNamespace("c", ChartNamespaceUri);
            return chart;
        }

        private void GenerateScatterChart(OfficeXmlNode chart) {
            uint xId = GenerateAxisId();
            uint yId = GenerateAxisId();

            OfficeXmlNode plotArea = PlotAreaOf(chart);
            OfficeXmlNode scatter = CreateScatterChart(plotArea, xId, yId);
            OfficeXmlNode xAxis = AddValueAxisInternal(plotArea, xId, yId, "b");
            OfficeXmlNode yAxis = AddValueAxisInternal(plotArea, yId, xId, "l");

            plotArea.Append(scatter);
            plotArea.Append(xAxis);
            plotArea.Append(yAxis);
        }

        private OfficeXmlNode AddScatterChartSeries(OfficeXmlNode group, uint index, string series, OfficeIMO.OfficeColor color, List<double> xValues, List<double> yValues) {
            OfficeXmlNode scatterSeries = NewSeries(group, index, series);
            scatterSeries.Append(AddShapeProperties(scatterSeries, color));
            AddScatterValues(AddC(scatterSeries, "xVal"), xValues);
            AddScatterValues(AddC(scatterSeries, "yVal"), yValues);
            return scatterSeries;
        }

        private static void AddScatterValues(OfficeXmlNode values, List<double> data) {
            OfficeXmlNode literal = AddC(values, "numLit");
            AddC(literal, "ptCount", (uint)data.Count);
            for (int i = 0; i < data.Count; i++) {
                // Ensure decimal values use invariant culture (period as decimal separator) for OpenXML compatibility
                string valueText = OfficeIMO.Core.Internal.OfficeRoundTripText.Format(data[i], System.Globalization.CultureInfo.InvariantCulture);
                AddPoint(literal, (uint)i, valueText);
            }
        }

        private static OfficeXmlNode CreateRadarChart(OfficeXmlNode plotArea, uint catAxisId, uint valAxisId) {
            OfficeXmlNode chart = NewC(plotArea, "radarChart");
            chart.DeclareNamespace("c", ChartNamespaceUri);
            AddC(chart, "radarStyle", "standard");
            chart.Append(AddDataLabel(chart));
            AddC(chart, "axId", catAxisId).DeclareNamespace("c", ChartNamespaceUri);
            AddC(chart, "axId", valAxisId).DeclareNamespace("c", ChartNamespaceUri);
            return chart;
        }

        private void GenerateRadarChart(OfficeXmlNode chart) {
            uint catId = GenerateAxisId();
            uint valId = GenerateAxisId();

            OfficeXmlNode plotArea = PlotAreaOf(chart);
            OfficeXmlNode radarChart = CreateRadarChart(plotArea, catId, valId);
            OfficeXmlNode catAxis = AddCategoryAxisInternal(plotArea, catId, valId, "b");
            OfficeXmlNode valAxis = AddValueAxisInternal(plotArea, valId, catId, "l");

            plotArea.Append(radarChart);
            plotArea.Append(catAxis);
            plotArea.Append(valAxis);
        }

        private OfficeXmlNode AddRadarChartSeries<T>(OfficeXmlNode group, uint index, string series, OfficeIMO.OfficeColor color, List<string> categories, List<T> values) {
            OfficeXmlNode radarSeries = NewSeries(group, index, series);
            radarSeries.Append(AddShapeProperties(radarSeries, color));
            radarSeries.Append(AddCategoryAxisData(radarSeries, categories));
            radarSeries.Append(AddValuesAxisData(radarSeries, values));
            return radarSeries;
        }

        private static OfficeXmlNode CreateBar3DChart(OfficeXmlNode plotArea, uint catAxisId, uint valAxisId) {
            OfficeXmlNode chart = NewC(plotArea, "bar3DChart");
            chart.DeclareNamespace("c", ChartNamespaceUri);

            // Append elements in correct OpenXML schema order for Bar3DChart:
            // 1. barDir, 2. grouping, 3. varyColors, 4. ser (added later via InsertSeries),
            // 5. dLbls, 6. gapWidth, 7. gapDepth, 8. shape, 9. axId, 10. extLst
            // gapWidth comes after the ser elements and is added with the first series.
            AddC(chart, "barDir", "col");
            AddC(chart, "grouping", "clustered");

            // Series are inserted before the axis ids.
            AddC(chart, "axId", catAxisId).DeclareNamespace("c", ChartNamespaceUri);
            AddC(chart, "axId", valAxisId).DeclareNamespace("c", ChartNamespaceUri);
            return chart;
        }

        private void GenerateBar3DChart(OfficeXmlNode chart) {
            uint catId = GenerateAxisId();
            uint valId = GenerateAxisId();

            OfficeXmlNode plotArea = PlotAreaOf(chart);
            OfficeXmlNode chart3d = CreateBar3DChart(plotArea, catId, valId);
            OfficeXmlNode catAxis = AddCategoryAxisInternal(plotArea, catId, valId, "b");
            OfficeXmlNode valAxis = AddValueAxisInternal(plotArea, valId, catId, "l");

            plotArea.Append(chart3d);
            plotArea.Append(catAxis);
            plotArea.Append(valAxis);
        }

        private static OfficeXmlNode CreateLine3DChart(OfficeXmlNode plotArea, uint catAxisId, uint valAxisId, uint seriesId) {
            OfficeXmlNode chart = NewC(plotArea, "line3DChart");
            chart.DeclareNamespace("c", ChartNamespaceUri);

            // Add elements in the correct OpenXML schema order for Line3DChart:
            // 1. grouping, 2. varyColors, 3. ser (LineChartSeries - inserted by InsertSeries), 4. dLbls, 5. dropLines, 6. gapDepth, 7. axId (3x), 8. extLst
            AddC(chart, "grouping", "standard");
            AddC(chart, "varyColors", false);

            // Elements that come after series - InsertSeries will place series between varyColors and dLbls
            chart.Append(AddDataLabel(chart));
            AddC(chart, "gapDepth", "150");

            // Line3DChart requires 3 axis IDs (CategoryAxis, ValueAxis, SeriesAxis)
            AddC(chart, "axId", catAxisId).DeclareNamespace("c", ChartNamespaceUri);
            AddC(chart, "axId", valAxisId).DeclareNamespace("c", ChartNamespaceUri);
            AddC(chart, "axId", seriesId).DeclareNamespace("c", ChartNamespaceUri);
            return chart;
        }

        private static OfficeXmlNode AddSeriesAxisInternal(OfficeXmlNode plotArea, uint axisId, uint crossAxisId, string position) {
            OfficeXmlNode axis = NewC(plotArea, "serAx");
            AddC(axis, "axId", axisId);
            AddC(AddC(axis, "scaling"), "orientation", "minMax");
            AddC(axis, "axPos", position);
            AddC(axis, "crossAx", crossAxisId);
            AddC(axis, "crosses", "autoZero");
            return axis;
        }

        private void GenerateLine3DChart(OfficeXmlNode chart) {
            uint catId = GenerateAxisId();
            uint valId = GenerateAxisId();
            uint seriesId = GenerateAxisId(); // Add third axis for Line3D

            OfficeXmlNode plotArea = PlotAreaOf(chart);
            OfficeXmlNode chart3d = CreateLine3DChart(plotArea, catId, valId, seriesId);
            OfficeXmlNode catAxis = AddCategoryAxisInternal(plotArea, catId, valId, "b");
            OfficeXmlNode valAxis = AddValueAxisInternal(plotArea, valId, catId, "l");
            OfficeXmlNode seriesAxis = AddSeriesAxisInternal(plotArea, seriesId, catId, "r");

            plotArea.Append(chart3d);
            plotArea.Append(catAxis);
            plotArea.Append(valAxis);
            plotArea.Append(seriesAxis);
        }

        private static void GeneratePie3DChart(OfficeXmlNode chart) {
            OfficeXmlNode plotArea = PlotAreaOf(chart);
            OfficeXmlNode pie3d = NewC(plotArea, "pie3DChart");
            pie3d.DeclareNamespace("c", ChartNamespaceUri);
            pie3d.Append(AddDataLabel(pie3d));
            plotArea.Append(pie3d);
        }

        private static OfficeXmlNode CreateArea3DChart(OfficeXmlNode plotArea, uint catAxisId, uint valAxisId) {
            OfficeXmlNode chart = NewC(plotArea, "area3DChart");
            chart.DeclareNamespace("c", ChartNamespaceUri);
            AddC(chart, "grouping", "standard");
            chart.Append(AddDataLabel(chart));
            AddC(chart, "gapDepth", "150");
            AddC(chart, "axId", catAxisId).DeclareNamespace("c", ChartNamespaceUri);
            AddC(chart, "axId", valAxisId).DeclareNamespace("c", ChartNamespaceUri);
            return chart;
        }

        private void GenerateArea3DChart(OfficeXmlNode chart) {
            uint catId = GenerateAxisId();
            uint valId = GenerateAxisId();

            OfficeXmlNode plotArea = PlotAreaOf(chart);
            OfficeXmlNode area3d = CreateArea3DChart(plotArea, catId, valId);
            OfficeXmlNode catAxis = AddCategoryAxisInternal(plotArea, catId, valId, "b");
            OfficeXmlNode valAxis = AddValueAxisInternal(plotArea, valId, catId, "l");

            plotArea.Append(area3d);
            plotArea.Append(catAxis);
            plotArea.Append(valAxis);
        }

        private void EnsureChartExistsScatter() {
            if (_chart.IsNone) CreateNewChart(GenerateScatterChart);
            else RejectOtherChartGroup("scatterChart", "scatter");
        }

        private void EnsureChartExistsRadar() {
            if (_chart.IsNone) CreateNewChart(GenerateRadarChart);
            else RejectOtherChartGroup("radarChart", "radar");
        }

        private void EnsureChartExistsArea3D() {
            if (_chart.IsNone) CreateNewChart(GenerateArea3DChart);
            else RejectOtherChartGroup("area3DChart", "3-D area");
        }

        private void EnsureChartExistsBar3D() {
            if (_chart.IsNone) CreateNewChart(GenerateBar3DChart);
            else RejectOtherChartGroup("bar3DChart", "3-D bar");
        }

        private void EnsureChartExistsPie3D() {
            if (_chart.IsNone) CreateNewChart(GeneratePie3DChart, axisTitles: false);
            else RejectOtherChartGroup("pie3DChart", "3-D pie");
        }

        private void EnsureChartExistsLine3D() {
            if (_chart.IsNone) CreateNewChart(GenerateLine3DChart, axisTitles: false);
            else RejectOtherChartGroup("line3DChart", "3-D line");
        }
    }
}
