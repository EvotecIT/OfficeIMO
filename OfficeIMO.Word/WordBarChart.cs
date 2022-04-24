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

            var oChart = GenerateChart();
            BarChartSeries barChartSeries2 = AddBarChartSeries(1, "USA", Color.AliceBlue);
            BarChartSeries barChartSeries3 = AddBarChartSeries(2, "Brazil", Color.Brown);
            BarChartSeries barChartSeries1 = AddBarChartSeries(0, "Poland", Color.Green);

            var barChart = oChart.PlotArea.GetFirstChild<BarChart>();
            barChart.Append(barChartSeries1);
            barChart.Append(barChartSeries2);
            barChart.Append(barChartSeries3);

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

        internal static StringReference AddBar(UInt32Value index, string series) {
            StringReference stringReference1 = new StringReference();

            Formula formula1 = new Formula();
            formula1.Text = "";

            NumericValue numericValue1 = new NumericValue();
            numericValue1.Text = series;

            StringPoint stringPoint1 = new StringPoint() { Index = index };

            stringPoint1.Append(numericValue1);

            StringCache stringCache1 = new StringCache();
            stringCache1.Append(stringPoint1);

            stringReference1.Append(formula1);
            stringReference1.Append(stringCache1);

            return stringReference1;

        }

        internal static NumericPoint AddNumericPoint(UInt32Value index, string value) {
            NumericPoint numericPoint4 = new NumericPoint() { Index = index };
            NumericValue numericValue9 = new NumericValue();
            numericValue9.Text = value;

            numericPoint4.Append(numericValue9);
            return numericPoint4;
        }

        internal static StringPoint AddStringPoint(UInt32Value index, string text) {
            StringPoint stringPoint2 = new StringPoint() { Index = index };
            NumericValue numericValue2 = new NumericValue();
            numericValue2.Text = text;

            stringPoint2.Append(numericValue2);
            return stringPoint2;
        }

        //internal static CategoryAxis AddCategoryAxisData() {

        //}

        //internal static Values AddValuesAxisData() {

        //}

        internal static BarChartSeries AddBarChartSeries(UInt32Value index, string series, SixLabors.ImageSharp.Color color) {
            BarChartSeries barChartSeries1 = new BarChartSeries();

            Index index1 = new Index() { Val = index };
            Order order1 = new Order() { Val = index };

            SeriesText seriesText1 = new SeriesText();

            var stringReference1 = AddBar(0, series);

            seriesText1.Append(stringReference1);


            InvertIfNegative invertIfNegative1 = new InvertIfNegative();

            CategoryAxisData categoryAxisData1 = new CategoryAxisData();


            var chartShapeProperties1 = AddShapeProperties(color);

            barChartSeries1.Append(index1);
            barChartSeries1.Append(order1);
            barChartSeries1.Append(seriesText1);
            barChartSeries1.Append(chartShapeProperties1);
            barChartSeries1.Append(invertIfNegative1);


            NumberReference numberReference1 = new NumberReference();


            StringReference stringReference2 = new StringReference();
            Formula formula2 = new Formula() { Text = "" };
            //formula2.Text = "";

            StringCache stringCache2 = new StringCache();
            // PointCount pointCount1 = new PointCount() { Val = (UInt32Value)4U };


            var stringPoint2 = AddStringPoint(0, "Food");
            var stringPoint3 = AddStringPoint(1, "Housing");
            var stringPoint4 = AddStringPoint(2, "Transportation");
            var stringPoint5 = AddStringPoint(3, "Health Care");

            //stringCache2.Append(pointCount1);
            stringCache2.Append(stringPoint2);
            stringCache2.Append(stringPoint3);
            stringCache2.Append(stringPoint4);
            stringCache2.Append(stringPoint5);

            stringReference2.Append(formula2);
            stringReference2.Append(stringCache2);

            categoryAxisData1.Append(stringReference2);


            Formula formula3 = new Formula() { Text = "" };

            NumberingCache numberingCache1 = new NumberingCache();
            FormatCode formatCode1 = new FormatCode() { Text = "General" };
            //PointCount pointCount2 = new PointCount() { Val = (UInt32Value)4U };

            var numericPoint1 = AddNumericPoint(0, "200");
            var numericPoint2 = AddNumericPoint(1, "80");
            var numericPoint3 = AddNumericPoint(2, "110");
            var numericPoint4 = AddNumericPoint(3, "60");
            //var numericPoint5 = AddNumericPoint(4, "500");

            numberingCache1.Append(formatCode1);
            // numberingCache1.Append(pointCount2);
            numberingCache1.Append(numericPoint1);
            numberingCache1.Append(numericPoint2);
            numberingCache1.Append(numericPoint3);
            numberingCache1.Append(numericPoint4);
            //numberingCache1.Append(numericPoint5);
            numberReference1.Append(formula3);
            numberReference1.Append(numberingCache1);

            Values values1 = new Values() { NumberReference = numberReference1 };
            //values1.Append(numberReference1);

            barChartSeries1.Append(categoryAxisData1);
            barChartSeries1.Append(values1);
            return barChartSeries1;
        }

        public static Chart GenerateChart() {
            Chart chart1 = new Chart();
            AutoTitleDeleted autoTitleDeleted1 = new AutoTitleDeleted() { Val = false };

            PlotArea plotArea1 = new PlotArea();
            Layout layout1 = new Layout();
            BarChart barChart1 = CreateBarChart();
            DataLabels dataLabels1 = AddDataLabel();
            barChart1.Append(dataLabels1);
            CategoryAxis categoryAxis1 = AddCategoryAxis();
            ValueAxis valueAxis1 = AddValueAxis();
            plotArea1.Append(layout1);
            plotArea1.Append(categoryAxis1);
            plotArea1.Append(valueAxis1);
            Legend legend1 = AddLegend();
            PlotVisibleOnly plotVisibleOnly1 = new PlotVisibleOnly() { Val = true };
            DisplayBlanksAs displayBlanksAs1 = new DisplayBlanksAs() { Val = DisplayBlanksAsValues.Gap };
            ShowDataLabelsOverMaximum showDataLabelsOverMaximum1 = new ShowDataLabelsOverMaximum() { Val = false };
            chart1.Append(autoTitleDeleted1);
            chart1.Append(legend1);
            chart1.Append(plotVisibleOnly1);
            chart1.Append(displayBlanksAs1);
            chart1.Append(showDataLabelsOverMaximum1);
            chart1.Append(plotArea1);

            //BarChartSeries barChartSeries1 = AddBarChartSeries(0, "Poland", Color.Green);
            //BarChartSeries barChartSeries2 = AddBarChartSeries(1, "USA", Color.AliceBlue);
            //BarChartSeries barChartSeries3 = AddBarChartSeries(2, "Brazil", Color.Brown);
            //barChart1.Append(barChartSeries1);
            //barChart1.Append(barChartSeries2);
            //barChart1.Append(barChartSeries3);
            plotArea1.Append(barChart1);
            return chart1;
        }
    }
}