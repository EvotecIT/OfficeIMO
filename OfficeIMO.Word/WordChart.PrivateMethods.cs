using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml;

namespace OfficeIMO.Word {
    public partial class WordChart {
        private CategoryAxisData InitializeCategoryAxisData() {
            var pieChartSeries = InitializePieChartSeries();
            CategoryAxisData categoryAxis = pieChartSeries?.GetFirstChild<CategoryAxisData>();
            // If CategoryAxisData does not exist, create it
            if (categoryAxis == null) {
                categoryAxis = new CategoryAxisData();
                StringLiteral stringLiteral = new StringLiteral();
                categoryAxis.Append(stringLiteral);
            }
            return categoryAxis;
        }

        private NumberLiteral InitializeNumberLiteral() {
            NumberLiteral literal = _chart?.PlotArea?.GetFirstChild<DocumentFormat.OpenXml.Drawing.Charts.ValueAxis>()?.GetFirstChild<Values>()?.GetFirstChild<NumberLiteral>();
            // If NumberLiteral does not exist, create it
            if (literal == null) {
                literal = new NumberLiteral();
                FormatCode format = new FormatCode() { Text = "General" };
                literal.Append(format);
            }
            return literal;
        }

        private Values InitializeValues() {
            var pieChartSeries = InitializePieChartSeries();
            Values values = pieChartSeries?.GetFirstChild<Values>() ?? new Values() { NumberLiteral = InitializeNumberLiteral() };
            return values;
        }

        private PieChartSeries InitializePieChartSeries() {
            if (_chart != null) {
                var pieChart = _chart.PlotArea.GetFirstChild<PieChart>();
                OpenXmlCompositeElement chartElement = pieChart;
                if (chartElement == null) {
                    chartElement = _chart.PlotArea.GetFirstChild<Pie3DChart>();
                }

                if (chartElement != null) {
                    var pieChartSeries = chartElement.GetFirstChild<PieChartSeries>();
                    var dataLabels = chartElement.GetFirstChild<DataLabels>() ?? chartElement.AppendChild(AddDataLabel());

                    if (pieChartSeries == null) {
                        pieChartSeries = CreatePieChartSeries(_index, "Title?");
                        chartElement.InsertBefore(pieChartSeries, dataLabels);
                    }
                    return pieChartSeries;
                }
            }
            return null;
        }

        private PieChartSeries CreatePieChartSeries(UInt32Value index, string series) {
            PieChartSeries pieChartSeries1 = new PieChartSeries();
            DocumentFormat.OpenXml.Drawing.Charts.Index index1 = new DocumentFormat.OpenXml.Drawing.Charts.Index() { Val = index };

            Order order1 = new Order() { Val = index };
            SeriesText seriesText1 = new SeriesText();

            var stringReference1 = AddSeries(0, series);
            seriesText1.Append(stringReference1);

            pieChartSeries1.Append(index1);
            pieChartSeries1.Append(order1);
            pieChartSeries1.Append(seriesText1);
            return pieChartSeries1;
        }

        /// <summary>
        /// Adds single category to charts
        /// </summary>
        /// <param name="category">The category.</param>
        private void AddSingleCategory(string category) {
            var pieChartSeries = InitializePieChartSeries();

            CategoryAxisData categoryAxis = InitializeCategoryAxisData();

            StringLiteral stringLiteral = categoryAxis.GetFirstChild<StringLiteral>();
            // If StringLiteral does not exist, create it
            if (stringLiteral == null) {
                stringLiteral = new StringLiteral();
                categoryAxis.Append(stringLiteral);
            }
            stringLiteral.Append(new StringPoint() { Index = _currentIndexCategory, NumericValue = new DocumentFormat.OpenXml.Drawing.Charts.NumericValue() { Text = category } });
            // Update the PointCount
            PointCount pointCount = stringLiteral.GetFirstChild<PointCount>();
            if (pointCount != null) {
                pointCount.Val = _currentIndexCategory + 1;
            } else {
                stringLiteral.InsertAt(new PointCount() { Val = 1 }, 0);
            }
            // Increment the current index
            _currentIndexCategory++;

            if (!pieChartSeries.Elements<CategoryAxisData>().Any()) {
                pieChartSeries.Append(categoryAxis);
            }
        }

        /// <summary>
        /// Adds the single value to charts
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="data">The data.</param>
        private void AddSingleValue<T>(T data) {
            // Initialize the PieChartSeries
            var pieChartSeries = InitializePieChartSeries();
            // Initialize the Values
            Values values = InitializeValues();
            // Initialize the NumberLiteral
            NumberLiteral literal = values.GetFirstChild<NumberLiteral>() ?? InitializeNumberLiteral();

            // Ensure decimal values use invariant culture (period as decimal separator) for OpenXML compatibility
            string valueText = data is double d ? d.ToString(System.Globalization.CultureInfo.InvariantCulture) :
                              data is float f ? f.ToString(System.Globalization.CultureInfo.InvariantCulture) :
                              data?.ToString() ?? "0";

            literal.Append(new NumericPoint() { Index = _currentIndexValues, NumericValue = new NumericValue() { Text = valueText } });
            // Update the PointCount
            PointCount pointCount = literal.GetFirstChild<PointCount>();
            if (pointCount != null) {
                pointCount.Val = _currentIndexValues + 1;
            } else {
                int pos = literal.Elements<FormatCode>().Any() ? 1 : 0;
                literal.InsertAt(new PointCount() { Val = 1 }, pos);
            }
            // Increment the current index
            _currentIndexValues++;
            // add values to the series if it does not exist
            if (!pieChartSeries.Elements<Values>().Any()) {
                pieChartSeries.Append(values);
            }
        }

        private Chart CreatePieChart(Chart chart) {
            PieChart pieChart1 = new PieChart();
            pieChart1.AddNamespaceDeclaration("c", "http://schemas.openxmlformats.org/drawingml/2006/chart");
            chart.PlotArea.Append(pieChart1);
            return chart;
        }

        private Chart GenerateChartBar(Chart chart) {
            UInt32Value catId = GenerateAxisId();
            UInt32Value valId = GenerateAxisId();

            BarChart barChart1 = CreateBarChart(catId, valId);
            CategoryAxis categoryAxis1 = AddCategoryAxisInternal(catId, valId, AxisPositionValues.Bottom);
            ValueAxis valueAxis1 = AddValueAxisInternal(valId, catId, AxisPositionValues.Left);
            chart.PlotArea.Append(barChart1);
            chart.PlotArea.Append(categoryAxis1);
            chart.PlotArea.Append(valueAxis1);

            return chart;
        }
        private BarChart CreateBarChart(UInt32Value catAxisId, UInt32Value valAxisId, BarDirectionValues? barDirection = null) {
            barDirection ??= BarDirectionValues.Bar;
            BarChart barChart1 = new BarChart();
            barChart1.AddNamespaceDeclaration("c", "http://schemas.openxmlformats.org/drawingml/2006/chart");

            BarDirection barDirection1 = new BarDirection() { Val = barDirection };
            BarGrouping barGrouping1 = new BarGrouping() { Val = BarGroupingValues.Standard };
            DataLabels dataLabels1 = AddDataLabel();
            GapWidth gapWidth1 = new GapWidth() { Val = (UInt16Value)200U };
            Overlap overlap1 = new Overlap() { Val = 0 };

            AxisId axisId1 = new AxisId() { Val = catAxisId };
            axisId1.AddNamespaceDeclaration("c", "http://schemas.openxmlformats.org/drawingml/2006/chart");

            AxisId axisId2 = new AxisId() { Val = valAxisId };
            axisId2.AddNamespaceDeclaration("c", "http://schemas.openxmlformats.org/drawingml/2006/chart");

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
            barChart1.Append(barDirection1);
            barChart1.Append(barGrouping1);
            barChart1.Append(dataLabels1);
            barChart1.Append(gapWidth1);
            barChart1.Append(overlap1);
            barChart1.Append(axisId1);
            barChart1.Append(axisId2);
            return barChart1;
        }

        private void EnsureChartExistsPie() {
            // minimum required to create chart
            if (_chart == null) {
                _chart = GenerateChart();
                _chart = CreatePieChart(_chart);
                _chartPart.ChartSpace.Append(_chart);

                // since the title may have changed, we need to update it
                UpdateTitle();
            }
        }

        private void EnsureChartExistsBar() {
            if (_chart == null) {
                _chart = GenerateChart();
                _chart = GenerateChartBar(_chart);
                _chartPart.ChartSpace.Append(_chart);

                // since the title may have changed, we need to update it
                UpdateTitle();
            }
        }

        private void EnsureChartExistsArea() {
            if (_chart == null) {
                _chart = GenerateChart();
                _chart = GenerateAreaChart(_chart);
                _chartPart.ChartSpace.Append(_chart);

                // since the title may have changed, we need to update it
                UpdateTitle();
            }
        }

        private void EnsureChartExistsLine() {
            if (_chart == null) {
                _chart = GenerateChart();
                _chart = GenerateLineChart(_chart);
                _chartPart.ChartSpace.Append(_chart);

                // since the title may have changed, we need to update it
                UpdateTitle();
            }
        }


        private BarChartSeries AddBarChartSeries<T>(UInt32Value index, string series, SixLabors.ImageSharp.Color color, List<string> categories, List<T> data) {
            BarChartSeries barChartSeries1 = new BarChartSeries();
            DocumentFormat.OpenXml.Drawing.Charts.Index index1 = new DocumentFormat.OpenXml.Drawing.Charts.Index() { Val = index };
            Order order1 = new Order() { Val = index };
            SeriesText seriesText1 = new SeriesText();
            var stringReference1 = AddSeries(0, series);
            seriesText1.Append(stringReference1);
            InvertIfNegative invertIfNegative1 = new InvertIfNegative();
            var chartShapeProperties1 = AddShapeProperties(color);
            barChartSeries1.Append(index1);
            barChartSeries1.Append(order1);
            barChartSeries1.Append(seriesText1);
            barChartSeries1.Append(chartShapeProperties1);
            barChartSeries1.Append(invertIfNegative1);
            Values values1 = AddValuesAxisData(data);
            CategoryAxisData categoryAxisData1 = AddCategoryAxisData(categories);
            barChartSeries1.Append(categoryAxisData1);
            barChartSeries1.Append(values1);
            return barChartSeries1;
        }

        private ChartShapeProperties AddShapeProperties(SixLabors.ImageSharp.Color color) {
            ChartShapeProperties chartShapeProperties1 = new ChartShapeProperties();

            DocumentFormat.OpenXml.Drawing.SolidFill solidFill1 = new DocumentFormat.OpenXml.Drawing.SolidFill();
            DocumentFormat.OpenXml.Drawing.RgbColorModelHex rgbColorModelHex1 = new DocumentFormat.OpenXml.Drawing.RgbColorModelHex() { Val = color.ToHexColor() };

            solidFill1.Append(rgbColorModelHex1);
            chartShapeProperties1.Append(solidFill1);

            return chartShapeProperties1;

        }
        private LineChart CreateLineChart(UInt32Value catAxisId, UInt32Value valAxisId) {
            LineChart lineChart1 = new LineChart();
            lineChart1.AddNamespaceDeclaration("c", "http://schemas.openxmlformats.org/drawingml/2006/chart");
            Grouping grouping1 = new Grouping() { Val = GroupingValues.Standard };

            DataLabels dataLabels1 = AddDataLabel();

            lineChart1.Append(grouping1);
            lineChart1.Append(dataLabels1);

            AxisId axisId1 = new AxisId() { Val = catAxisId };
            axisId1.AddNamespaceDeclaration("c", "http://schemas.openxmlformats.org/drawingml/2006/chart");

            AxisId axisId2 = new AxisId() { Val = valAxisId };
            axisId2.AddNamespaceDeclaration("c", "http://schemas.openxmlformats.org/drawingml/2006/chart");

            lineChart1.Append(axisId1);
            lineChart1.Append(axisId2);
            return lineChart1;
        }

        private Chart GenerateLineChart(Chart chart) {
            UInt32Value catId = GenerateAxisId();
            UInt32Value valId = GenerateAxisId();

            LineChart lineChart1 = CreateLineChart(catId, valId);
            CategoryAxis categoryAxis1 = AddCategoryAxisInternal(catId, valId, AxisPositionValues.Bottom);
            ValueAxis valueAxis1 = AddValueAxisInternal(valId, catId, AxisPositionValues.Left);
            //chart.PlotArea.Append(layout1);
            chart.PlotArea.Append(lineChart1);
            chart.PlotArea.Append(categoryAxis1);
            chart.PlotArea.Append(valueAxis1);
            return chart;
        }
        private LineChartSeries AddLineChartSeries<T>(UInt32Value index, string series, SixLabors.ImageSharp.Color color, List<string> categories, List<T> data) {
            LineChartSeries lineChartSeries1 = new LineChartSeries();
            DocumentFormat.OpenXml.Drawing.Charts.Index index1 = new DocumentFormat.OpenXml.Drawing.Charts.Index() { Val = index };
            Order order1 = new Order() { Val = index };

            SeriesText seriesText1 = new SeriesText();

            var stringReference1 = AddSeries(0, series);

            seriesText1.Append(stringReference1);

            // Note: InvertIfNegative is not valid for LineChartSeries according to OpenXML schema
            // It's only valid for BarChartSeries and other chart types, but not LineChartSeries

            var chartShapeProperties1 = AddShapeProperties(color);

            Values values1 = AddValuesAxisData(data);
            CategoryAxisData categoryAxisData1 = AddCategoryAxisData(categories);

            // Append elements in correct OpenXML schema order for LineChartSeries:
            // 1. idx, 2. order, 3. tx (seriesText), 4. spPr (chartShapeProperties),
            // 5. marker, 6. pictureOptions, 7. dPt, 8. dLbls, 9. trendline, 10. errBars,
            // 11. cat (categoryAxisData), 12. val (values), 13. smooth, 14. extLst
            lineChartSeries1.Append(index1);
            lineChartSeries1.Append(order1);
            lineChartSeries1.Append(seriesText1);
            lineChartSeries1.Append(chartShapeProperties1);
            lineChartSeries1.Append(categoryAxisData1);
            lineChartSeries1.Append(values1);

            return lineChartSeries1;


        }

        private AreaChart CreateAreaChart(UInt32Value catAxisId, UInt32Value valAxisId) {
            AreaChart chart = new AreaChart();
            chart.AddNamespaceDeclaration("c", "http://schemas.openxmlformats.org/drawingml/2006/chart");
            Grouping grouping1 = new Grouping() { Val = GroupingValues.Standard };

            chart.Append(grouping1);

            DataLabels dataLabels1 = AddDataLabel();
            chart.Append(dataLabels1);

            AxisId axisId1 = new AxisId() { Val = catAxisId };
            axisId1.AddNamespaceDeclaration("c", "http://schemas.openxmlformats.org/drawingml/2006/chart");

            AxisId axisId2 = new AxisId() { Val = valAxisId };
            axisId2.AddNamespaceDeclaration("c", "http://schemas.openxmlformats.org/drawingml/2006/chart");

            chart.Append(axisId1);
            chart.Append(axisId2);
            return chart;
        }

        private Chart GenerateAreaChart(Chart chart) {
            UInt32Value catId = GenerateAxisId();
            UInt32Value valId = GenerateAxisId();

            AreaChart areaChart = CreateAreaChart(catId, valId);

            CategoryAxis categoryAxis1 = AddCategoryAxisInternal(catId, valId, AxisPositionValues.Bottom);
            ValueAxis valueAxis1 = AddValueAxisInternal(valId, catId, AxisPositionValues.Left);

            //chart.PlotArea.Append(layout1);
            chart.PlotArea.Append(areaChart);
            chart.PlotArea.Append(categoryAxis1);
            chart.PlotArea.Append(valueAxis1);


            return chart;
        }

        private AreaChartSeries AddAreaChartSeries<T>(UInt32Value index, string series, SixLabors.ImageSharp.Color color, List<string> categories, List<T> data) {
            AreaChartSeries lineChartSeries1 = new AreaChartSeries();
            DocumentFormat.OpenXml.Drawing.Charts.Index index1 = new DocumentFormat.OpenXml.Drawing.Charts.Index() { Val = index };
            Order order1 = new Order() { Val = index };

            SeriesText seriesText1 = new SeriesText();


            NumericValue numericValue1 = new NumericValue();
            numericValue1.Text = series;

            seriesText1.Append(numericValue1);


            var chartShapeProperties1 = AddShapeProperties(color);

            Values values1 = AddValuesAxisData(data);
            CategoryAxisData categoryAxisData1 = AddCategoryAxisData(categories);


            lineChartSeries1.Append(index1);
            lineChartSeries1.Append(order1);
            lineChartSeries1.Append(seriesText1);
            lineChartSeries1.Append(chartShapeProperties1);

            lineChartSeries1.Append(categoryAxisData1);
            lineChartSeries1.Append(values1);

            return lineChartSeries1;
        }

        private static UInt32Value GenerateAxisId() {
            int id = System.Threading.Interlocked.Increment(ref _axisIdSeed);
            return (UInt32Value)(uint)id;
        }

        private static void InsertSeries(OpenXmlCompositeElement chartElement, OpenXmlCompositeElement series) {
            var dataLabels = chartElement.GetFirstChild<DataLabels>();
            if (dataLabels != null) {
                chartElement.InsertBefore(series, dataLabels);
                return;
            }

            var axis = chartElement.Elements<AxisId>().FirstOrDefault();
            if (axis != null) {
                chartElement.InsertBefore(series, axis);
            } else {
                chartElement.Append(series);
            }
        }

        private ScatterChart CreateScatterChart(UInt32Value xAxisId, UInt32Value yAxisId) {
            ScatterChart chart = new ScatterChart();
            chart.AddNamespaceDeclaration("c", "http://schemas.openxmlformats.org/drawingml/2006/chart");
            chart.Append(new ScatterStyle() { Val = ScatterStyleValues.Marker });

            DataLabels labels = AddDataLabel();
            chart.Append(labels);

            AxisId axisId1 = new AxisId() { Val = xAxisId };
            axisId1.AddNamespaceDeclaration("c", "http://schemas.openxmlformats.org/drawingml/2006/chart");

            AxisId axisId2 = new AxisId() { Val = yAxisId };
            axisId2.AddNamespaceDeclaration("c", "http://schemas.openxmlformats.org/drawingml/2006/chart");

            chart.Append(axisId1);
            chart.Append(axisId2);
            return chart;
        }

        private Chart GenerateScatterChart(Chart chart) {
            UInt32Value xId = GenerateAxisId();
            UInt32Value yId = GenerateAxisId();

            ScatterChart scatter = CreateScatterChart(xId, yId);

            ValueAxis xAxis = AddValueAxisInternal(xId, yId, AxisPositionValues.Bottom);
            ValueAxis yAxis = AddValueAxisInternal(yId, xId, AxisPositionValues.Left);

            chart.PlotArea.Append(scatter);
            chart.PlotArea.Append(xAxis);
            chart.PlotArea.Append(yAxis);

            return chart;
        }

        private ScatterChartSeries AddScatterChartSeries(UInt32Value index, string series, SixLabors.ImageSharp.Color color, List<double> xValues, List<double> yValues) {
            ScatterChartSeries scSeries = new ScatterChartSeries();
            DocumentFormat.OpenXml.Drawing.Charts.Index idx = new DocumentFormat.OpenXml.Drawing.Charts.Index() { Val = index };
            Order order = new Order() { Val = index };

            SeriesText text = new SeriesText();
            var seriesRef = AddSeries(0, series);
            text.Append(seriesRef);

            var shape = AddShapeProperties(color); XValues x = new XValues();
            NumberLiteral xLit = new NumberLiteral();
            xLit.Append(new PointCount() { Val = (uint)xValues.Count }); for (int i = 0; i < xValues.Count; i++) {
                // Ensure decimal values use invariant culture (period as decimal separator) for OpenXML compatibility
                string xValueText = xValues[i].ToString(System.Globalization.CultureInfo.InvariantCulture);
                xLit.Append(new NumericPoint() { Index = (uint)i, NumericValue = new NumericValue(xValueText) });
            }
            x.Append(xLit);

            YValues y = new YValues();
            NumberLiteral yLit = new NumberLiteral();
            yLit.Append(new PointCount() { Val = (uint)yValues.Count }); for (int i = 0; i < yValues.Count; i++) {
                // Ensure decimal values use invariant culture (period as decimal separator) for OpenXML compatibility
                string yValueText = yValues[i].ToString(System.Globalization.CultureInfo.InvariantCulture);
                yLit.Append(new NumericPoint() { Index = (uint)i, NumericValue = new NumericValue(yValueText) });
            }
            y.Append(yLit);

            scSeries.Append(idx);
            scSeries.Append(order);
            scSeries.Append(text);
            scSeries.Append(shape);
            scSeries.Append(x);
            scSeries.Append(y);

            return scSeries;
        }

        private RadarChart CreateRadarChart(UInt32Value catAxisId, UInt32Value valAxisId) {
            RadarChart chart = new RadarChart();
            chart.AddNamespaceDeclaration("c", "http://schemas.openxmlformats.org/drawingml/2006/chart");

            RadarStyle style = new RadarStyle() { Val = RadarStyleValues.Standard };
            chart.Append(style);

            DataLabels labels = AddDataLabel();
            chart.Append(labels);

            AxisId axisId1 = new AxisId() { Val = catAxisId };
            axisId1.AddNamespaceDeclaration("c", "http://schemas.openxmlformats.org/drawingml/2006/chart");

            AxisId axisId2 = new AxisId() { Val = valAxisId };
            axisId2.AddNamespaceDeclaration("c", "http://schemas.openxmlformats.org/drawingml/2006/chart");

            chart.Append(axisId1);
            chart.Append(axisId2);
            return chart;
        }

        private Chart GenerateRadarChart(Chart chart) {
            UInt32Value catId = GenerateAxisId();
            UInt32Value valId = GenerateAxisId();

            RadarChart radarChart = CreateRadarChart(catId, valId);
            CategoryAxis catAxis = AddCategoryAxisInternal(catId, valId, AxisPositionValues.Bottom);
            ValueAxis valAxis = AddValueAxisInternal(valId, catId, AxisPositionValues.Left);

            chart.PlotArea.Append(radarChart);
            chart.PlotArea.Append(catAxis);
            chart.PlotArea.Append(valAxis);
            return chart;
        }

        private RadarChartSeries AddRadarChartSeries<T>(UInt32Value index, string series, SixLabors.ImageSharp.Color color, List<string> categories, List<T> values) {
            RadarChartSeries radarSeries = new RadarChartSeries();
            DocumentFormat.OpenXml.Drawing.Charts.Index idx = new DocumentFormat.OpenXml.Drawing.Charts.Index() { Val = index };
            Order order = new Order() { Val = index };

            SeriesText text = new SeriesText();
            var seriesRef = AddSeries(0, series);
            text.Append(seriesRef);

            var shape = AddShapeProperties(color);
            CategoryAxisData cats = AddCategoryAxisData(categories);
            Values vals = AddValuesAxisData(values);

            radarSeries.Append(idx);
            radarSeries.Append(order);
            radarSeries.Append(text);
            radarSeries.Append(shape);
            radarSeries.Append(cats);
            radarSeries.Append(vals);
            return radarSeries;
        }
        private Bar3DChart CreateBar3DChart(UInt32Value catAxisId, UInt32Value valAxisId) {
            Bar3DChart chart = new Bar3DChart();
            chart.AddNamespaceDeclaration("c", "http://schemas.openxmlformats.org/drawingml/2006/chart");

            // Append elements in correct OpenXML schema order for Bar3DChart:
            // 1. barDir, 2. grouping, 3. varyColors, 4. ser (added later via InsertSeries),
            // 5. dLbls, 6. gapWidth, 7. gapDepth, 8. shape, 9. axId, 10. extLst
            // Note: ser elements must come before gapWidth, which must come before axId
            chart.Append(new BarDirection() { Val = BarDirectionValues.Column });
            chart.Append(new BarGrouping() { Val = BarGroupingValues.Clustered });

            // Don't add gapWidth here - it should come after ser elements
            // We'll add it in a different location or modify InsertSeries to handle this properly

            // Add axId elements at the end - series will be inserted before these by InsertSeries method
            AxisId axisId1 = new AxisId() { Val = catAxisId };
            axisId1.AddNamespaceDeclaration("c", "http://schemas.openxmlformats.org/drawingml/2006/chart");
            AxisId axisId2 = new AxisId() { Val = valAxisId };
            axisId2.AddNamespaceDeclaration("c", "http://schemas.openxmlformats.org/drawingml/2006/chart");

            chart.Append(axisId1);
            chart.Append(axisId2);
            return chart;
        }

        private Chart GenerateBar3DChart(Chart chart) {
            UInt32Value catId = GenerateAxisId();
            UInt32Value valId = GenerateAxisId();

            Bar3DChart chart3d = CreateBar3DChart(catId, valId);
            CategoryAxis catAxis = AddCategoryAxisInternal(catId, valId, AxisPositionValues.Bottom);
            ValueAxis valAxis = AddValueAxisInternal(valId, catId, AxisPositionValues.Left);

            chart.PlotArea.Append(chart3d);
            chart.PlotArea.Append(catAxis);
            chart.PlotArea.Append(valAxis);
            return chart;
        }

        private BarChartSeries AddBar3DChartSeries<T>(UInt32Value index, string series, SixLabors.ImageSharp.Color color, List<string> categories, List<T> values) {
            BarChartSeries series3d = AddBarChartSeries(index, series, color, categories, values);
            return series3d;
        }

        private Line3DChart CreateLine3DChart(UInt32Value catAxisId, UInt32Value valAxisId) {
            Line3DChart chart = new Line3DChart();
            chart.AddNamespaceDeclaration("c", "http://schemas.openxmlformats.org/drawingml/2006/chart");

            Grouping grouping = new Grouping() { Val = GroupingValues.Standard };
            DataLabels labels = AddDataLabel();
            GapDepth gapDepth = new GapDepth() { Val = (UInt16Value)150U };

            chart.Append(grouping);
            chart.Append(labels);
            chart.Append(gapDepth);

            AxisId axisId1 = new AxisId() { Val = catAxisId };
            axisId1.AddNamespaceDeclaration("c", "http://schemas.openxmlformats.org/drawingml/2006/chart");
            AxisId axisId2 = new AxisId() { Val = valAxisId };
            axisId2.AddNamespaceDeclaration("c", "http://schemas.openxmlformats.org/drawingml/2006/chart");

            chart.Append(axisId1);
            chart.Append(axisId2);
            return chart;
        }

        private Chart GenerateLine3DChart(Chart chart) {
            UInt32Value catId = GenerateAxisId();
            UInt32Value valId = GenerateAxisId();

            Line3DChart chart3d = CreateLine3DChart(catId, valId);
            CategoryAxis catAxis = AddCategoryAxisInternal(catId, valId, AxisPositionValues.Bottom);
            ValueAxis valAxis = AddValueAxisInternal(valId, catId, AxisPositionValues.Left);

            chart.PlotArea.Append(chart3d);
            chart.PlotArea.Append(catAxis);
            chart.PlotArea.Append(valAxis);
            return chart;
        }

        private LineChartSeries AddLine3DChartSeries<T>(UInt32Value index, string series, SixLabors.ImageSharp.Color color, List<string> categories, List<T> values) {
            LineChartSeries series3d = AddLineChartSeries(index, series, color, categories, values);
            return series3d;
        }

        private PieChartSeries AddPie3DChartSeries<T>(UInt32Value index, string series, SixLabors.ImageSharp.Color color, List<string> categories, List<T> values) {
            PieChartSeries pieSeries = new PieChartSeries();
            DocumentFormat.OpenXml.Drawing.Charts.Index idx = new DocumentFormat.OpenXml.Drawing.Charts.Index() { Val = index };
            Order order = new Order() { Val = index };

            SeriesText text = new SeriesText();
            var seriesRef = AddSeries(0, series);
            text.Append(seriesRef);

            var shape = AddShapeProperties(color);
            CategoryAxisData cats = AddCategoryAxisData(categories);
            Values vals = AddValuesAxisData(values);

            pieSeries.Append(idx);
            pieSeries.Append(order);
            pieSeries.Append(text);
            pieSeries.Append(shape);
            pieSeries.Append(cats);
            pieSeries.Append(vals);
            return pieSeries;
        }

        private Pie3DChart CreatePie3DChart() {
            Pie3DChart chart3d = new Pie3DChart();
            chart3d.AddNamespaceDeclaration("c", "http://schemas.openxmlformats.org/drawingml/2006/chart");

            DataLabels labels = AddDataLabel();
            chart3d.Append(labels);
            return chart3d;
        }

        private Chart GeneratePie3DChart(Chart chart) {
            Pie3DChart pie3d = CreatePie3DChart();
            chart.PlotArea.Append(pie3d);
            return chart;
        }

        private Area3DChart CreateArea3DChart(UInt32Value catAxisId, UInt32Value valAxisId) {
            Area3DChart chart = new Area3DChart();
            chart.AddNamespaceDeclaration("c", "http://schemas.openxmlformats.org/drawingml/2006/chart");

            Grouping grouping = new Grouping() { Val = GroupingValues.Standard };
            chart.Append(grouping);

            DataLabels labels = AddDataLabel();
            GapDepth gapDepth = new GapDepth() { Val = (UInt16Value)150U };

            chart.Append(labels);
            chart.Append(gapDepth);

            AxisId axisId1 = new AxisId() { Val = catAxisId };
            axisId1.AddNamespaceDeclaration("c", "http://schemas.openxmlformats.org/drawingml/2006/chart");
            AxisId axisId2 = new AxisId() { Val = valAxisId };
            axisId2.AddNamespaceDeclaration("c", "http://schemas.openxmlformats.org/drawingml/2006/chart");

            chart.Append(axisId1);
            chart.Append(axisId2);

            return chart;
        }

        private Chart GenerateArea3DChart(Chart chart) {
            UInt32Value catId = GenerateAxisId();
            UInt32Value valId = GenerateAxisId();

            Area3DChart area3d = CreateArea3DChart(catId, valId);
            CategoryAxis catAxis = AddCategoryAxisInternal(catId, valId, AxisPositionValues.Bottom);
            ValueAxis valAxis = AddValueAxisInternal(valId, catId, AxisPositionValues.Left);

            chart.PlotArea.Append(area3d);
            chart.PlotArea.Append(catAxis);
            chart.PlotArea.Append(valAxis);

            return chart;
        }

        private AreaChartSeries AddArea3DChartSeries<T>(UInt32Value index, string series, SixLabors.ImageSharp.Color color, List<string> categories, List<T> values) {
            return AddAreaChartSeries(index, series, color, categories, values);
        }

        private void EnsureChartExistsScatter() {
            if (_chart == null) {
                _chart = GenerateChart();
                _chart = GenerateScatterChart(_chart);
                _chartPart.ChartSpace.Append(_chart);
                UpdateTitle();
            }
        }

        private void EnsureChartExistsRadar() {
            if (_chart == null) {
                _chart = GenerateChart();
                _chart = GenerateRadarChart(_chart);
                _chartPart.ChartSpace.Append(_chart);
                UpdateTitle();
            }
        }

        private void EnsureChartExistsArea3D() {
            if (_chart == null) {
                _chart = GenerateChart();
                _chart = GenerateArea3DChart(_chart);
                _chartPart.ChartSpace.Append(_chart);
                UpdateTitle();
            }
        }

        private void EnsureChartExistsBar3D() {
            if (_chart == null) {
                _chart = GenerateChart();
                _chart = GenerateBar3DChart(_chart);
                _chartPart.ChartSpace.Append(_chart);
                UpdateTitle();
            }
        }

        private void EnsureChartExistsPie3D() {
            if (_chart == null) {
                _chart = GenerateChart();
                _chart = GeneratePie3DChart(_chart);
                _chartPart.ChartSpace.Append(_chart);
                UpdateTitle();
            }
        }

        private void EnsureChartExistsLine3D() {
            if (_chart == null) {
                _chart = GenerateChart();
                _chart = GenerateLine3DChart(_chart);
                _chartPart.ChartSpace.Append(_chart);
                UpdateTitle();
            }
        }
    }
}
