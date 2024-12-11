using DocumentFormat.OpenXml.Drawing.Charts;

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
                if (pieChart != null) {
                    var pieChartSeries = pieChart.GetFirstChild<PieChartSeries>();
                    if (pieChartSeries == null) {
                        pieChartSeries = CreatePieChartSeries(_index, "Title?");
                        pieChart.Append(pieChartSeries);

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

            InvertIfNegative invertIfNegative1 = new InvertIfNegative();
            pieChartSeries1.Append(index1);
            pieChartSeries1.Append(order1);
            pieChartSeries1.Append(seriesText1);
            pieChartSeries1.Append(invertIfNegative1);
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
            literal.Append(new NumericPoint() { Index = _currentIndexValues, NumericValue = new NumericValue() { Text = data.ToString() } });
            // Update the PointCount
            PointCount pointCount = literal.GetFirstChild<PointCount>();
            if (pointCount != null) {
                pointCount.Val = _currentIndexValues + 1;
            } else {
                literal.InsertAt(new PointCount() { Val = 1 }, 0);
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
            DataLabels dataLabels1 = AddDataLabel();
            pieChart1.AddNamespaceDeclaration("c", "http://schemas.openxmlformats.org/drawingml/2006/chart");
            pieChart1.Append(dataLabels1);
            chart.PlotArea.Append(pieChart1);
            return chart;
        }

        private Chart GenerateChartBar(Chart chart) {
            BarChart barChart1 = CreateBarChart();
            CategoryAxis categoryAxis1 = AddCategoryAxis();
            ValueAxis valueAxis1 = AddValueAxis();
            chart.PlotArea.Append(categoryAxis1);
            chart.PlotArea.Append(valueAxis1);
            chart.PlotArea.Append(barChart1);

            return chart;
        }

        private BarChart CreateBarChart(BarDirectionValues? barDirection = null) {
            barDirection ??= BarDirectionValues.Bar;
            BarChart barChart1 = new BarChart();
            barChart1.AddNamespaceDeclaration("c", "http://schemas.openxmlformats.org/drawingml/2006/chart");

            DataLabels dataLabels1 = AddDataLabel();
            barChart1.Append(dataLabels1);

            BarDirection barDirection1 = new BarDirection() { Val = barDirection };
            BarGrouping barGrouping1 = new BarGrouping() { Val = BarGroupingValues.Standard };
            GapWidth gapWidth1 = new GapWidth() { Val = (UInt16Value)200U };

            AxisId axisId1 = new AxisId() { Val = (UInt32Value)148921728U };
            axisId1.AddNamespaceDeclaration("c", "http://schemas.openxmlformats.org/drawingml/2006/chart");

            AxisId axisId2 = new AxisId() { Val = (UInt32Value)154227840U };
            axisId2.AddNamespaceDeclaration("c", "http://schemas.openxmlformats.org/drawingml/2006/chart");
            Overlap overlap1 = new Overlap() { Val = 0 };

            barChart1.Append(barDirection1);
            barChart1.Append(barGrouping1);
            barChart1.Append(gapWidth1);
            barChart1.Append(axisId1);
            barChart1.Append(axisId2);
            barChart1.Append(overlap1);
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
        private LineChart CreateLineChart() {
            LineChart lineChart1 = new LineChart();
            lineChart1.AddNamespaceDeclaration("c", "http://schemas.openxmlformats.org/drawingml/2006/chart");
            Grouping grouping1 = new Grouping() { Val = GroupingValues.Standard };

            DataLabels dataLabels1 = AddDataLabel();

            lineChart1.Append(grouping1);
            lineChart1.Append(dataLabels1);

            AxisId axisId1 = new AxisId() { Val = (UInt32Value)148921728U };
            axisId1.AddNamespaceDeclaration("c", "http://schemas.openxmlformats.org/drawingml/2006/chart");

            AxisId axisId2 = new AxisId() { Val = (UInt32Value)154227840U };
            axisId2.AddNamespaceDeclaration("c", "http://schemas.openxmlformats.org/drawingml/2006/chart");

            lineChart1.Append(axisId1);
            lineChart1.Append(axisId2);
            return lineChart1;
        }

        private Chart GenerateLineChart(Chart chart) {
            LineChart lineChart1 = CreateLineChart();
            CategoryAxis categoryAxis1 = AddCategoryAxis();
            ValueAxis valueAxis1 = AddValueAxis();
            //chart.PlotArea.Append(layout1);
            chart.PlotArea.Append(categoryAxis1);
            chart.PlotArea.Append(valueAxis1);
            chart.PlotArea.Append(lineChart1);
            return chart;
        }

        private LineChartSeries AddLineChartSeries<T>(UInt32Value index, string series, SixLabors.ImageSharp.Color color, List<string> categories, List<T> data) {
            LineChartSeries lineChartSeries1 = new LineChartSeries();
            DocumentFormat.OpenXml.Drawing.Charts.Index index1 = new DocumentFormat.OpenXml.Drawing.Charts.Index() { Val = index };
            Order order1 = new Order() { Val = index };

            SeriesText seriesText1 = new SeriesText();

            var stringReference1 = AddSeries(0, series);

            seriesText1.Append(stringReference1);

            InvertIfNegative invertIfNegative1 = new InvertIfNegative();

            var chartShapeProperties1 = AddShapeProperties(color);

            Values values1 = AddValuesAxisData(data);
            CategoryAxisData categoryAxisData1 = AddCategoryAxisData(categories);


            lineChartSeries1.Append(index1);
            lineChartSeries1.Append(order1);
            lineChartSeries1.Append(seriesText1);
            lineChartSeries1.Append(chartShapeProperties1);
            lineChartSeries1.Append(invertIfNegative1);
            lineChartSeries1.Append(categoryAxisData1);
            lineChartSeries1.Append(values1);

            return lineChartSeries1;


        }

        private AreaChart CreateAreaChart() {
            AreaChart chart = new AreaChart();
            chart.AddNamespaceDeclaration("c", "http://schemas.openxmlformats.org/drawingml/2006/chart");
            Grouping grouping1 = new Grouping() { Val = GroupingValues.Standard };

            DataLabels dataLabels1 = AddDataLabel();
            chart.Append(dataLabels1);

            chart.Append(grouping1);

            AxisId axisId1 = new AxisId() { Val = (UInt32Value)148921728U };
            axisId1.AddNamespaceDeclaration("c", "http://schemas.openxmlformats.org/drawingml/2006/chart");

            AxisId axisId2 = new AxisId() { Val = (UInt32Value)154227840U };
            axisId2.AddNamespaceDeclaration("c", "http://schemas.openxmlformats.org/drawingml/2006/chart");

            chart.Append(axisId1);
            chart.Append(axisId2);
            return chart;
        }

        private Chart GenerateAreaChart(Chart chart) {
            AreaChart areaChart = CreateAreaChart();

            CategoryAxis categoryAxis1 = AddCategoryAxis();
            ValueAxis valueAxis1 = AddValueAxis();

            //chart.PlotArea.Append(layout1);
            chart.PlotArea.Append(categoryAxis1);
            chart.PlotArea.Append(valueAxis1);
            chart.PlotArea.Append(areaChart);


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
    }
}
