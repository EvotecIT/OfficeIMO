using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    public class WordLineChart : WordChart {

        public static WordChart AddLineChart(WordDocument wordDocument, WordParagraph paragraph, bool roundedCorners = false) {
            _document = wordDocument;
            _paragraph = paragraph;

            // minimum required to create chart
            var oChart = GenerateChart();
            oChart = GenerateLineChart(oChart);

            // this is data for bar chart
            List<string> categories = new List<string>() {
                "Food", "Housing", "Mix", "Data"
            };

            LineChartSeries lineChartSeries1 = AddLineChartSeries(1, "USA", SixLabors.ImageSharp.Color.AliceBlue, categories, new List<object>() { 15, 20, 30, 150 });
            LineChartSeries lineChartSeries2 = AddLineChartSeries(2, "Brazil", SixLabors.ImageSharp.Color.Brown, categories, new List<object>() { 20, 20, 300, 150 });
            LineChartSeries lineChartSeries3 = AddLineChartSeries(0, "Poland", SixLabors.ImageSharp.Color.Green, categories, new List<object>() { 13, 20, 230, 150 });

            var lineChart = oChart.PlotArea.GetFirstChild<LineChart>();
            lineChart.Append(lineChartSeries1);
            lineChart.Append(lineChartSeries2);
            lineChart.Append(lineChartSeries3);

            // inserts chart into document
            InsertChart(wordDocument, paragraph, oChart, roundedCorners);

            return new WordChart();
        }

        internal static LineChart CreateLineChart() {
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

        private static Chart GenerateLineChart(Chart chart) {
            LineChart lineChart1 = CreateLineChart();

            CategoryAxis categoryAxis1 = AddCategoryAxis();
            ValueAxis valueAxis1 = AddValueAxis();



            //chart.PlotArea.Append(layout1);
            chart.PlotArea.Append(categoryAxis1);
            chart.PlotArea.Append(valueAxis1);
            chart.PlotArea.Append(lineChart1);


            return chart;
        }

        internal static LineChartSeries AddLineChartSeries(UInt32Value index, string series, SixLabors.ImageSharp.Color color, List<string> categories, List<object> data) {
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

        internal static ChartShapeProperties AddShapeProperties(SixLabors.ImageSharp.Color color) {
            ChartShapeProperties chartShapeProperties1 = new ChartShapeProperties();

            DocumentFormat.OpenXml.Drawing.Outline outline1 = new DocumentFormat.OpenXml.Drawing.Outline();

            DocumentFormat.OpenXml.Drawing.SolidFill solidFill1 = new DocumentFormat.OpenXml.Drawing.SolidFill();
            DocumentFormat.OpenXml.Drawing.RgbColorModelHex rgbColorModelHex1 = new DocumentFormat.OpenXml.Drawing.RgbColorModelHex() { Val = color.ToHexColor() };

            solidFill1.Append(rgbColorModelHex1);

            outline1.Append(solidFill1);

            chartShapeProperties1.Append(outline1);

            return chartShapeProperties1;
        }
    }
}
