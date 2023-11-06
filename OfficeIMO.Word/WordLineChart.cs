using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    public class WordLineChart : WordChart {

        public static WordChart AddLineChart(WordDocument wordDocument, WordParagraph paragraph, string title = null, bool roundedCorners = false, int width = 600, int height = 600) {
            _document = wordDocument;
            _paragraph = paragraph;

            // minimum required to create chart
            var oChart = GenerateChart(title);
            oChart = GenerateLineChart(oChart);

            // inserts chart into document
            InsertChart(wordDocument, paragraph, oChart, roundedCorners, width, height);

            var drawing = paragraph._paragraph.OfType<Drawing>().FirstOrDefault();

            return new WordChart(_document, _paragraph._paragraph, drawing);
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

        internal static LineChartSeries AddLineChartSeries<T>(UInt32Value index, string series, SixLabors.ImageSharp.Color color, List<string> categories, List<T> data) {
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

        public WordLineChart(WordDocument document, Paragraph paragraph, Drawing drawing) : base(document, paragraph, drawing) {
        }
    }
}
