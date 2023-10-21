using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Color = SixLabors.ImageSharp.Color;

namespace OfficeIMO.Word {
    public class WordPieChart : WordChart {
        public static WordChart AddPieChart(WordDocument wordDocument, WordParagraph paragraph, string title = null, bool roundedCorners = false, int width = 600, int height = 600) {
            _document = wordDocument;
            _paragraph = paragraph;

            // minimum required to create chart
            var oChart = GenerateChart(title);
            oChart = CreatePieChart(oChart);

            // inserts chart into document
            InsertChart(wordDocument, paragraph, oChart, roundedCorners, width, height);

            var drawing = paragraph._paragraph.OfType<Drawing>().FirstOrDefault();

            return new WordChart(_document, _paragraph._paragraph, drawing);
        }

        internal static Chart CreatePieChart(Chart chart) {
            PieChart pieChart1 = new PieChart();
            DataLabels dataLabels1 = AddDataLabel();
            pieChart1.AddNamespaceDeclaration("c", "http://schemas.openxmlformats.org/drawingml/2006/chart");
            pieChart1.Append(dataLabels1);


            chart.PlotArea.Append(pieChart1);
            return chart;
        }

        internal static PieChartSeries AddPieChartSeries<T>(UInt32Value index, string series, List<string> categories, List<T> data) {
            PieChartSeries pieChartSeries1 = new PieChartSeries();

            DocumentFormat.OpenXml.Drawing.Charts.Index index1 = new DocumentFormat.OpenXml.Drawing.Charts.Index() { Val = index };
            Order order1 = new Order() { Val = index };


            SeriesText seriesText1 = new SeriesText();

            var stringReference1 = AddSeries(0, series);
            seriesText1.Append(stringReference1);

            InvertIfNegative invertIfNegative1 = new InvertIfNegative();
            CategoryAxisData categoryAxisData1 = AddCategoryAxisData(categories);
            Values values1 = AddValuesAxisData(data);

            pieChartSeries1.Append(index1);
            pieChartSeries1.Append(order1);
            pieChartSeries1.Append(seriesText1);
            pieChartSeries1.Append(invertIfNegative1);
            pieChartSeries1.Append(values1);
            pieChartSeries1.Append(categoryAxisData1);
            return pieChartSeries1;
        }

        public WordPieChart(WordDocument document, Paragraph paragraph, Drawing drawing) : base(document, paragraph, drawing) {
        }
    }
}
