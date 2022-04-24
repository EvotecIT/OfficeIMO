using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Office2013.Drawing.ChartStyle;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using CategoryAxis = DocumentFormat.OpenXml.Drawing.Charts.CategoryAxis;
using Color = SixLabors.ImageSharp.Color;
using DataLabels = DocumentFormat.OpenXml.Drawing.Charts.DataLabels;
using Legend = DocumentFormat.OpenXml.Drawing.Charts.Legend;
using PlotArea = DocumentFormat.OpenXml.Drawing.Charts.PlotArea;
using ValueAxis = DocumentFormat.OpenXml.Drawing.Charts.ValueAxis;

namespace OfficeIMO.Word {
    public class WordBarChart : WordChart {
        public static WordChart AddBarChart(WordDocument wordDocument, WordParagraph paragraph, bool roundedCorners = false) {
            _document = wordDocument;
            _paragraph = paragraph;

            // minimum required to create chart
            var oChart = GenerateChart();
            oChart = GenerateChartBar(oChart);

            // this is data for bar chart
            List<string> categories = new List<string>() {
                "Food", "Housing", "Mix", "Data"
            };

            BarChartSeries barChartSeries2 = AddBarChartSeries(1, "USA", Color.AliceBlue, categories, new List<object>() { 15, 20, 30, 150 });
            BarChartSeries barChartSeries3 = AddBarChartSeries(2, "Brazil", Color.Brown, categories, new List<object>() { 20, 20, 300, 150 });
            BarChartSeries barChartSeries1 = AddBarChartSeries(0, "Poland", Color.Green, categories, new List<object>() { 13, 20, 230, 150 });

            var barChart = oChart.PlotArea.GetFirstChild<BarChart>();
            barChart.Append(barChartSeries1);
            barChart.Append(barChartSeries2);
            barChart.Append(barChartSeries3);

            // inserts chart into document
            InsertChart(wordDocument, paragraph, oChart, roundedCorners);

            return new WordChart();
        }

        internal static BarChart CreateBarChart(BarDirectionValues barDirection = BarDirectionValues.Bar) {
            BarChart barChart1 = new BarChart();
            barChart1.AddNamespaceDeclaration("c", "http://schemas.openxmlformats.org/drawingml/2006/chart");

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

        internal static ChartShapeProperties AddShapeProperties(SixLabors.ImageSharp.Color color) {
            ChartShapeProperties chartShapeProperties1 = new ChartShapeProperties();

            DocumentFormat.OpenXml.Drawing.SolidFill solidFill1 = new DocumentFormat.OpenXml.Drawing.SolidFill();
            DocumentFormat.OpenXml.Drawing.RgbColorModelHex rgbColorModelHex1 = new DocumentFormat.OpenXml.Drawing.RgbColorModelHex() { Val = color.ToHexColor() };

            solidFill1.Append(rgbColorModelHex1);
            chartShapeProperties1.Append(solidFill1);

            return chartShapeProperties1;

        }








        internal static BarChartSeries AddBarChartSeries(UInt32Value index, string series, SixLabors.ImageSharp.Color color, List<string> categories, List<object> data) {
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

        private static Chart GenerateChartBar(Chart chart) {
            BarChart barChart1 = CreateBarChart();
            DataLabels dataLabels1 = AddDataLabel();
            barChart1.Append(dataLabels1);

            CategoryAxis categoryAxis1 = AddCategoryAxis();
            ValueAxis valueAxis1 = AddValueAxis();

            chart.PlotArea.Append(categoryAxis1);
            chart.PlotArea.Append(valueAxis1);
            chart.PlotArea.Append(barChart1);
            return chart;
        }
    }
}