using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Color = SixLabors.ImageSharp.Color;

namespace OfficeIMO.Word {
    public class WordPieChart : WordChart {
        public static WordChart AddPieChart(WordDocument wordDocument, WordParagraph paragraph, bool roundedCorners = false) {
            _document = wordDocument;
            _paragraph = paragraph;

            // minimum required to create chart
            var oChart = GenerateChart();
            oChart = CreatePieChart(oChart);

            //// this is data for pie chart
            //List<string> categories = new List<string>() {
            //    "Food", "Housing", "Mix", "Data"
            //};

            //PieChartSeries pieChartSeries1 = AddPieChartSeries(0, "USA", Color.AliceBlue, categories, new List<int>() { 15, 20, 30, 150 });

            //var pieChart = oChart.PlotArea.GetFirstChild<PieChart>();
            //pieChart.Append(pieChartSeries1);

            // inserts chart into document
            InsertChart(wordDocument, paragraph, oChart, roundedCorners);

            return new WordChart();
        }

        internal static Chart CreatePieChart(Chart chart) {
            PieChart pieChart1 = new PieChart();
            DataLabels dataLabels1 = AddDataLabel();
            pieChart1.AddNamespaceDeclaration("c", "http://schemas.openxmlformats.org/drawingml/2006/chart");
            pieChart1.Append(dataLabels1);


            chart.PlotArea.Append(pieChart1);
            return chart;
        }

        internal static PieChartSeries AddPieChartSeries(UInt32Value index, string series, SixLabors.ImageSharp.Color color, List<string> categories, List<int> data) {
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

        //public static Chart GenerateChart1() {
        //    Chart chart1 = new Chart();
        //    AutoTitleDeleted autoTitleDeleted1 = new AutoTitleDeleted() { Val = false };

        //    PlotArea plotArea1 = new PlotArea();
        //    Layout layout1 = new Layout();

        //    //PieChart pieChart1 = new PieChart();
        //    //pieChart1.AddNamespaceDeclaration("c", "http://schemas.openxmlformats.org/drawingml/2006/chart");

        //    //DataLabels dataLabels1 = AddDataLabel();

        //    //PieChartSeries pieChartSeries1 = new PieChartSeries();
        //    //DocumentFormat.OpenXml.Drawing.Charts.Index index1 = new DocumentFormat.OpenXml.Drawing.Charts.Index() { Val = (UInt32Value)1U };
        //    //Order order1 = new Order() { Val = (UInt32Value)1U };

        //    //SeriesText seriesText1 = new SeriesText();

        //    //StringReference stringReference1 = new StringReference();
        //    //Formula formula1 = new Formula();
        //    //formula1.Text = "";

        //    //StringCache stringCache1 = new StringCache();

        //    //StringPoint stringPoint1 = new StringPoint() { Index = (UInt32Value)0U };
        //    //NumericValue numericValue1 = new NumericValue();
        //    //numericValue1.Text = "Brazil";

        //    //stringPoint1.Append(numericValue1);

        //    //stringCache1.Append(stringPoint1);

        //    //stringReference1.Append(formula1);
        //    //stringReference1.Append(stringCache1);

        //    //seriesText1.Append(stringReference1);
        //    //OpenXmlUnknownElement openXmlUnknownElement1 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<c:invertIfNegative xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\">0</c:invertIfNegative>");

        //    //CategoryAxisData categoryAxisData1 = new CategoryAxisData();

        //    //StringReference stringReference2 = new StringReference();
        //    //Formula formula2 = new Formula();
        //    //formula2.Text = "";

        //    //StringCache stringCache2 = new StringCache();
        //    //PointCount pointCount1 = new PointCount() { Val = (UInt32Value)4U };

        //    //StringPoint stringPoint2 = new StringPoint() { Index = (UInt32Value)0U };
        //    //NumericValue numericValue2 = new NumericValue();
        //    //numericValue2.Text = "Food";

        //    //stringPoint2.Append(numericValue2);

        //    //StringPoint stringPoint3 = new StringPoint() { Index = (UInt32Value)1U };
        //    //NumericValue numericValue3 = new NumericValue();
        //    //numericValue3.Text = "Housing";

        //    //stringPoint3.Append(numericValue3);

        //    //StringPoint stringPoint4 = new StringPoint() { Index = (UInt32Value)2U };
        //    //NumericValue numericValue4 = new NumericValue();
        //    //numericValue4.Text = "Transportation";

        //    //stringPoint4.Append(numericValue4);

        //    //StringPoint stringPoint5 = new StringPoint() { Index = (UInt32Value)3U };
        //    //NumericValue numericValue5 = new NumericValue();
        //    //numericValue5.Text = "Health Care";

        //    //stringPoint5.Append(numericValue5);

        //    //stringCache2.Append(pointCount1);
        //    //stringCache2.Append(stringPoint2);
        //    //stringCache2.Append(stringPoint3);
        //    //stringCache2.Append(stringPoint4);
        //    //stringCache2.Append(stringPoint5);

        //    //stringReference2.Append(formula2);
        //    //stringReference2.Append(stringCache2);

        //    //categoryAxisData1.Append(stringReference2);

        //    Values values1 = new Values();

        //    NumberReference numberReference1 = new NumberReference();
        //    Formula formula3 = new Formula();
        //    formula3.Text = "";

        //    NumberingCache numberingCache1 = new NumberingCache();
        //    FormatCode formatCode1 = new FormatCode();
        //    formatCode1.Text = "General";
        //    PointCount pointCount2 = new PointCount() { Val = (UInt32Value)4U };

        //    NumericPoint numericPoint1 = new NumericPoint() { Index = (UInt32Value)0U };
        //    NumericValue numericValue6 = new NumericValue();
        //    numericValue6.Text = "125";

        //    numericPoint1.Append(numericValue6);

        //    NumericPoint numericPoint2 = new NumericPoint() { Index = (UInt32Value)1U };
        //    NumericValue numericValue7 = new NumericValue();
        //    numericValue7.Text = "80";

        //    numericPoint2.Append(numericValue7);

        //    NumericPoint numericPoint3 = new NumericPoint() { Index = (UInt32Value)2U };
        //    NumericValue numericValue8 = new NumericValue();
        //    numericValue8.Text = "110";

        //    numericPoint3.Append(numericValue8);

        //    NumericPoint numericPoint4 = new NumericPoint() { Index = (UInt32Value)3U };
        //    NumericValue numericValue9 = new NumericValue();
        //    numericValue9.Text = "60";

        //    numericPoint4.Append(numericValue9);

        //    numberingCache1.Append(formatCode1);
        //    numberingCache1.Append(pointCount2);
        //    numberingCache1.Append(numericPoint1);
        //    numberingCache1.Append(numericPoint2);
        //    numberingCache1.Append(numericPoint3);
        //    numberingCache1.Append(numericPoint4);

        //    numberReference1.Append(formula3);
        //    numberReference1.Append(numberingCache1);

        //    values1.Append(numberReference1);

        //    //pieChartSeries1.Append(index1);
        //    //pieChartSeries1.Append(order1);
        //    //pieChartSeries1.Append(seriesText1);
        //    //pieChartSeries1.Append(openXmlUnknownElement1);
        //    //pieChartSeries1.Append(categoryAxisData1);
        //    //pieChartSeries1.Append(values1);

        //    //pieChart1.Append(dataLabels1);
        //    //pieChart1.Append(pieChartSeries1);

        //    //plotArea1.Append(layout1);
        //    //plotArea1.Append(pieChart1);

        //    Legend legend1 = AddLegend();

        //    PlotVisibleOnly plotVisibleOnly1 = new PlotVisibleOnly() { Val = true };
        //    DisplayBlanksAs displayBlanksAs1 = new DisplayBlanksAs() { Val = DisplayBlanksAsValues.Gap };
        //    ShowDataLabelsOverMaximum showDataLabelsOverMaximum1 = new ShowDataLabelsOverMaximum() { Val = false };

        //    chart1.Append(autoTitleDeleted1);
        //    chart1.Append(plotArea1);
        //    chart1.Append(legend1);
        //    chart1.Append(plotVisibleOnly1);
        //    chart1.Append(displayBlanksAs1);
        //    chart1.Append(showDataLabelsOverMaximum1);
        //    return chart1;
        //}

        //public Run GenerateRun() {
        //    Run run1 = new Run();
        //    run1.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

        //    Drawing drawing1 = new Drawing();

        //    DocumentFormat.OpenXml.Drawing.Wordprocessing.Inline inline1 = new DocumentFormat.OpenXml.Drawing.Wordprocessing.Inline();
        //    inline1.AddNamespaceDeclaration("wp", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing");
        //    DocumentFormat.OpenXml.Drawing.Wordprocessing.Extent extent1 = new DocumentFormat.OpenXml.Drawing.Wordprocessing.Extent() { Cx = 5486400L, Cy = 3200400L };
        //    DocumentFormat.OpenXml.Drawing.Wordprocessing.EffectExtent effectExtent1 = new DocumentFormat.OpenXml.Drawing.Wordprocessing.EffectExtent() { LeftEdge = 0L, TopEdge = 0L, RightEdge = 19050L, BottomEdge = 19050L };
        //    DocumentFormat.OpenXml.Drawing.Wordprocessing.DocProperties docProperties1 = new DocumentFormat.OpenXml.Drawing.Wordprocessing.DocProperties() { Id = (UInt32Value)2U, Name = "chart" };

        //    DocumentFormat.OpenXml.Drawing.Graphic graphic1 = new DocumentFormat.OpenXml.Drawing.Graphic();
        //    graphic1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

        //    DocumentFormat.OpenXml.Drawing.GraphicData graphicData1 = new DocumentFormat.OpenXml.Drawing.GraphicData() { Uri = "http://schemas.openxmlformats.org/drawingml/2006/chart" };

        //    DocumentFormat.OpenXml.Drawing.Charts.ChartReference chartReference1 = new DocumentFormat.OpenXml.Drawing.Charts.ChartReference() { Id = "rId1" };
        //    chartReference1.AddNamespaceDeclaration("p6", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
        //    chartReference1.AddNamespaceDeclaration("c", "http://schemas.openxmlformats.org/drawingml/2006/chart");

        //    graphicData1.Append(chartReference1);

        //    graphic1.Append(graphicData1);

        //    inline1.Append(extent1);
        //    inline1.Append(effectExtent1);
        //    inline1.Append(docProperties1);
        //    inline1.Append(graphic1);

        //    drawing1.Append(inline1);

        //    run1.Append(drawing1);
        //    return run1;
        //}
    }
}
