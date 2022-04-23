using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    public class WordBarChart : WordChart {
        public static WordChart AddBarChart(WordDocument wordDocument, WordParagraph paragraph, bool roundedCorners = false) {
            _document = wordDocument;
            _paragraph = paragraph;

            var oChart = GenerateChart();
            InsertChart(wordDocument, paragraph, oChart, roundedCorners);

            return new WordChart();
        }

        public static Chart GenerateChart() {
            Chart chart1 = new Chart();
            AutoTitleDeleted autoTitleDeleted1 = new AutoTitleDeleted() { Val = false };

            PlotArea plotArea1 = new PlotArea();
            Layout layout1 = new Layout();

            BarChart barChart1 = new BarChart();
            barChart1.AddNamespaceDeclaration("c", "http://schemas.openxmlformats.org/drawingml/2006/chart");
            BarDirection barDirection1 = new BarDirection() { Val = BarDirectionValues.Bar };
            BarGrouping barGrouping1 = new BarGrouping() { Val = BarGroupingValues.Standard };
            GapWidth gapWidth1 = new GapWidth() { Val = (UInt16Value)200U };

            AxisId axisId1 = new AxisId() { Val = (UInt32Value)148921728U };
            axisId1.AddNamespaceDeclaration("c", "http://schemas.openxmlformats.org/drawingml/2006/chart");

            AxisId axisId2 = new AxisId() { Val = (UInt32Value)154227840U };
            axisId2.AddNamespaceDeclaration("c", "http://schemas.openxmlformats.org/drawingml/2006/chart");
            Overlap overlap1 = new Overlap() { Val = 0 };

            DataLabels dataLabels1 = AddDataLabel();

            BarChartSeries barChartSeries1 = new BarChartSeries();
            Index index1 = new Index() { Val = (UInt32Value)1U };
            Order order1 = new Order() { Val = (UInt32Value)1U };

            SeriesText seriesText1 = new SeriesText();

            StringReference stringReference1 = new StringReference();
            Formula formula1 = new Formula();
            formula1.Text = "";

            StringCache stringCache1 = new StringCache();

            StringPoint stringPoint1 = new StringPoint() { Index = (UInt32Value)0U };
            NumericValue numericValue1 = new NumericValue();
            numericValue1.Text = "Brazil";

            stringPoint1.Append(numericValue1);

            stringCache1.Append(stringPoint1);

            stringReference1.Append(formula1);
            stringReference1.Append(stringCache1);

            seriesText1.Append(stringReference1);

            ChartShapeProperties chartShapeProperties1 = new ChartShapeProperties();

            DocumentFormat.OpenXml.Drawing.SolidFill solidFill1 = new DocumentFormat.OpenXml.Drawing.SolidFill();
            DocumentFormat.OpenXml.Drawing.RgbColorModelHex rgbColorModelHex1 = new DocumentFormat.OpenXml.Drawing.RgbColorModelHex() { Val = "ADFF2F" };

            solidFill1.Append(rgbColorModelHex1);

            chartShapeProperties1.Append(solidFill1);
            InvertIfNegative invertIfNegative1 = new InvertIfNegative();

            CategoryAxisData categoryAxisData1 = new CategoryAxisData();

            StringReference stringReference2 = new StringReference();
            Formula formula2 = new Formula();
            formula2.Text = "";

            StringCache stringCache2 = new StringCache();
            PointCount pointCount1 = new PointCount() { Val = (UInt32Value)4U };

            StringPoint stringPoint2 = new StringPoint() { Index = (UInt32Value)0U };
            NumericValue numericValue2 = new NumericValue();
            numericValue2.Text = "Food";

            stringPoint2.Append(numericValue2);

            StringPoint stringPoint3 = new StringPoint() { Index = (UInt32Value)1U };
            NumericValue numericValue3 = new NumericValue();
            numericValue3.Text = "Housing";

            stringPoint3.Append(numericValue3);

            StringPoint stringPoint4 = new StringPoint() { Index = (UInt32Value)2U };
            NumericValue numericValue4 = new NumericValue();
            numericValue4.Text = "Transportation";

            stringPoint4.Append(numericValue4);

            StringPoint stringPoint5 = new StringPoint() { Index = (UInt32Value)3U };
            NumericValue numericValue5 = new NumericValue();
            numericValue5.Text = "Health Care";

            stringPoint5.Append(numericValue5);

            stringCache2.Append(pointCount1);
            stringCache2.Append(stringPoint2);
            stringCache2.Append(stringPoint3);
            stringCache2.Append(stringPoint4);
            stringCache2.Append(stringPoint5);

            stringReference2.Append(formula2);
            stringReference2.Append(stringCache2);

            categoryAxisData1.Append(stringReference2);

            Values values1 = new Values();

            NumberReference numberReference1 = new NumberReference();
            Formula formula3 = new Formula();
            formula3.Text = "";

            NumberingCache numberingCache1 = new NumberingCache();
            FormatCode formatCode1 = new FormatCode();
            formatCode1.Text = "General";
            PointCount pointCount2 = new PointCount() { Val = (UInt32Value)4U };

            NumericPoint numericPoint1 = new NumericPoint() { Index = (UInt32Value)0U };
            NumericValue numericValue6 = new NumericValue();
            numericValue6.Text = "125";

            numericPoint1.Append(numericValue6);

            NumericPoint numericPoint2 = new NumericPoint() { Index = (UInt32Value)1U };
            NumericValue numericValue7 = new NumericValue();
            numericValue7.Text = "80";

            numericPoint2.Append(numericValue7);

            NumericPoint numericPoint3 = new NumericPoint() { Index = (UInt32Value)2U };
            NumericValue numericValue8 = new NumericValue();
            numericValue8.Text = "110";

            numericPoint3.Append(numericValue8);

            NumericPoint numericPoint4 = new NumericPoint() { Index = (UInt32Value)3U };
            NumericValue numericValue9 = new NumericValue();
            numericValue9.Text = "60";

            numericPoint4.Append(numericValue9);

            numberingCache1.Append(formatCode1);
            numberingCache1.Append(pointCount2);
            numberingCache1.Append(numericPoint1);
            numberingCache1.Append(numericPoint2);
            numberingCache1.Append(numericPoint3);
            numberingCache1.Append(numericPoint4);

            numberReference1.Append(formula3);
            numberReference1.Append(numberingCache1);

            values1.Append(numberReference1);

            barChartSeries1.Append(index1);
            barChartSeries1.Append(order1);
            barChartSeries1.Append(seriesText1);
            barChartSeries1.Append(chartShapeProperties1);
            barChartSeries1.Append(invertIfNegative1);
            barChartSeries1.Append(categoryAxisData1);
            barChartSeries1.Append(values1);

            BarChartSeries barChartSeries2 = new BarChartSeries();
            Index index2 = new Index() { Val = (UInt32Value)2U };
            Order order2 = new Order() { Val = (UInt32Value)2U };

            SeriesText seriesText2 = new SeriesText();

            StringReference stringReference3 = new StringReference();
            Formula formula4 = new Formula();
            formula4.Text = "";

            StringCache stringCache3 = new StringCache();

            StringPoint stringPoint6 = new StringPoint() { Index = (UInt32Value)0U };
            NumericValue numericValue10 = new NumericValue();
            numericValue10.Text = "USA";

            stringPoint6.Append(numericValue10);

            stringCache3.Append(stringPoint6);

            stringReference3.Append(formula4);
            stringReference3.Append(stringCache3);

            seriesText2.Append(stringReference3);

            ChartShapeProperties chartShapeProperties2 = new ChartShapeProperties();

            DocumentFormat.OpenXml.Drawing.SolidFill solidFill2 = new DocumentFormat.OpenXml.Drawing.SolidFill();
            DocumentFormat.OpenXml.Drawing.RgbColorModelHex rgbColorModelHex2 = new DocumentFormat.OpenXml.Drawing.RgbColorModelHex() { Val = "ADD8E6" };

            solidFill2.Append(rgbColorModelHex2);

            chartShapeProperties2.Append(solidFill2);
            InvertIfNegative invertIfNegative2 = new InvertIfNegative();

            CategoryAxisData categoryAxisData2 = new CategoryAxisData();

            StringReference stringReference4 = new StringReference();
            Formula formula5 = new Formula();
            formula5.Text = "";

            StringCache stringCache4 = new StringCache();
            PointCount pointCount3 = new PointCount() { Val = (UInt32Value)4U };

            StringPoint stringPoint7 = new StringPoint() { Index = (UInt32Value)0U };
            NumericValue numericValue11 = new NumericValue();
            numericValue11.Text = "Food";

            stringPoint7.Append(numericValue11);

            StringPoint stringPoint8 = new StringPoint() { Index = (UInt32Value)1U };
            NumericValue numericValue12 = new NumericValue();
            numericValue12.Text = "Housing";

            stringPoint8.Append(numericValue12);

            StringPoint stringPoint9 = new StringPoint() { Index = (UInt32Value)2U };
            NumericValue numericValue13 = new NumericValue();
            numericValue13.Text = "Transportation";

            stringPoint9.Append(numericValue13);

            StringPoint stringPoint10 = new StringPoint() { Index = (UInt32Value)3U };
            NumericValue numericValue14 = new NumericValue();
            numericValue14.Text = "Health Care";

            stringPoint10.Append(numericValue14);

            stringCache4.Append(pointCount3);
            stringCache4.Append(stringPoint7);
            stringCache4.Append(stringPoint8);
            stringCache4.Append(stringPoint9);
            stringCache4.Append(stringPoint10);

            stringReference4.Append(formula5);
            stringReference4.Append(stringCache4);

            categoryAxisData2.Append(stringReference4);

            Values values2 = new Values();

            NumberReference numberReference2 = new NumberReference();
            Formula formula6 = new Formula();
            formula6.Text = "";

            NumberingCache numberingCache2 = new NumberingCache();
            FormatCode formatCode2 = new FormatCode();
            formatCode2.Text = "General";
            PointCount pointCount4 = new PointCount() { Val = (UInt32Value)4U };

            NumericPoint numericPoint5 = new NumericPoint() { Index = (UInt32Value)0U };
            NumericValue numericValue15 = new NumericValue();
            numericValue15.Text = "200";

            numericPoint5.Append(numericValue15);

            NumericPoint numericPoint6 = new NumericPoint() { Index = (UInt32Value)1U };
            NumericValue numericValue16 = new NumericValue();
            numericValue16.Text = "150";

            numericPoint6.Append(numericValue16);

            NumericPoint numericPoint7 = new NumericPoint() { Index = (UInt32Value)2U };
            NumericValue numericValue17 = new NumericValue();
            numericValue17.Text = "110";

            numericPoint7.Append(numericValue17);

            NumericPoint numericPoint8 = new NumericPoint() { Index = (UInt32Value)3U };
            NumericValue numericValue18 = new NumericValue();
            numericValue18.Text = "100";

            numericPoint8.Append(numericValue18);

            numberingCache2.Append(formatCode2);
            numberingCache2.Append(pointCount4);
            numberingCache2.Append(numericPoint5);
            numberingCache2.Append(numericPoint6);
            numberingCache2.Append(numericPoint7);
            numberingCache2.Append(numericPoint8);

            numberReference2.Append(formula6);
            numberReference2.Append(numberingCache2);

            values2.Append(numberReference2);

            barChartSeries2.Append(index2);
            barChartSeries2.Append(order2);
            barChartSeries2.Append(seriesText2);
            barChartSeries2.Append(chartShapeProperties2);
            barChartSeries2.Append(invertIfNegative2);
            barChartSeries2.Append(categoryAxisData2);
            barChartSeries2.Append(values2);

            BarChartSeries barChartSeries3 = new BarChartSeries();
            Index index3 = new Index() { Val = (UInt32Value)3U };
            Order order3 = new Order() { Val = (UInt32Value)3U };

            SeriesText seriesText3 = new SeriesText();

            StringReference stringReference5 = new StringReference();
            Formula formula7 = new Formula();
            formula7.Text = "";

            StringCache stringCache5 = new StringCache();

            StringPoint stringPoint11 = new StringPoint() { Index = (UInt32Value)0U };
            NumericValue numericValue19 = new NumericValue();
            numericValue19.Text = "Canada";

            stringPoint11.Append(numericValue19);

            stringCache5.Append(stringPoint11);

            stringReference5.Append(formula7);
            stringReference5.Append(stringCache5);

            seriesText3.Append(stringReference5);

            ChartShapeProperties chartShapeProperties3 = new ChartShapeProperties();

            DocumentFormat.OpenXml.Drawing.SolidFill solidFill3 = new DocumentFormat.OpenXml.Drawing.SolidFill();
            DocumentFormat.OpenXml.Drawing.RgbColorModelHex rgbColorModelHex3 = new DocumentFormat.OpenXml.Drawing.RgbColorModelHex() { Val = "808080" };

            solidFill3.Append(rgbColorModelHex3);

            chartShapeProperties3.Append(solidFill3);
            InvertIfNegative invertIfNegative3 = new InvertIfNegative();

            CategoryAxisData categoryAxisData3 = new CategoryAxisData();

            StringReference stringReference6 = new StringReference();
            Formula formula8 = new Formula();
            formula8.Text = "";

            StringCache stringCache6 = new StringCache();
            PointCount pointCount5 = new PointCount() { Val = (UInt32Value)4U };

            StringPoint stringPoint12 = new StringPoint() { Index = (UInt32Value)0U };
            NumericValue numericValue20 = new NumericValue();
            numericValue20.Text = "Food";

            stringPoint12.Append(numericValue20);

            StringPoint stringPoint13 = new StringPoint() { Index = (UInt32Value)1U };
            NumericValue numericValue21 = new NumericValue();
            numericValue21.Text = "Housing";

            stringPoint13.Append(numericValue21);

            StringPoint stringPoint14 = new StringPoint() { Index = (UInt32Value)2U };
            NumericValue numericValue22 = new NumericValue();
            numericValue22.Text = "Transportation";

            stringPoint14.Append(numericValue22);

            StringPoint stringPoint15 = new StringPoint() { Index = (UInt32Value)3U };
            NumericValue numericValue23 = new NumericValue();
            numericValue23.Text = "Health Care";

            stringPoint15.Append(numericValue23);

            stringCache6.Append(pointCount5);
            stringCache6.Append(stringPoint12);
            stringCache6.Append(stringPoint13);
            stringCache6.Append(stringPoint14);
            stringCache6.Append(stringPoint15);

            stringReference6.Append(formula8);
            stringReference6.Append(stringCache6);

            categoryAxisData3.Append(stringReference6);

            Values values3 = new Values();

            NumberReference numberReference3 = new NumberReference();
            Formula formula9 = new Formula();
            formula9.Text = "";

            NumberingCache numberingCache3 = new NumberingCache();
            FormatCode formatCode3 = new FormatCode();
            formatCode3.Text = "General";
            PointCount pointCount6 = new PointCount() { Val = (UInt32Value)4U };

            NumericPoint numericPoint9 = new NumericPoint() { Index = (UInt32Value)0U };
            NumericValue numericValue24 = new NumericValue();
            numericValue24.Text = "100";

            numericPoint9.Append(numericValue24);

            NumericPoint numericPoint10 = new NumericPoint() { Index = (UInt32Value)1U };
            NumericValue numericValue25 = new NumericValue();
            numericValue25.Text = "120";

            numericPoint10.Append(numericValue25);

            NumericPoint numericPoint11 = new NumericPoint() { Index = (UInt32Value)2U };
            NumericValue numericValue26 = new NumericValue();
            numericValue26.Text = "140";

            numericPoint11.Append(numericValue26);

            NumericPoint numericPoint12 = new NumericPoint() { Index = (UInt32Value)3U };
            NumericValue numericValue27 = new NumericValue();
            numericValue27.Text = "150";

            numericPoint12.Append(numericValue27);

            numberingCache3.Append(formatCode3);
            numberingCache3.Append(pointCount6);
            numberingCache3.Append(numericPoint9);
            numberingCache3.Append(numericPoint10);
            numberingCache3.Append(numericPoint11);
            numberingCache3.Append(numericPoint12);

            numberReference3.Append(formula9);
            numberReference3.Append(numberingCache3);

            values3.Append(numberReference3);

            barChartSeries3.Append(index3);
            barChartSeries3.Append(order3);
            barChartSeries3.Append(seriesText3);
            barChartSeries3.Append(chartShapeProperties3);
            barChartSeries3.Append(invertIfNegative3);
            barChartSeries3.Append(categoryAxisData3);
            barChartSeries3.Append(values3);

            barChart1.Append(barDirection1);
            barChart1.Append(barGrouping1);
            barChart1.Append(gapWidth1);
            barChart1.Append(axisId1);
            barChart1.Append(axisId2);
            barChart1.Append(overlap1);
            barChart1.Append(dataLabels1);
            barChart1.Append(barChartSeries1);
            barChart1.Append(barChartSeries2);
            barChart1.Append(barChartSeries3);

            CategoryAxis categoryAxis1 = AddCategoryAxis();

            ValueAxis valueAxis1 = AddValueAxis();

            plotArea1.Append(layout1);
            plotArea1.Append(barChart1);
            plotArea1.Append(categoryAxis1);
            plotArea1.Append(valueAxis1);

            Legend legend1 = AddLegend();

            PlotVisibleOnly plotVisibleOnly1 = new PlotVisibleOnly() { Val = true };
            DisplayBlanksAs displayBlanksAs1 = new DisplayBlanksAs() { Val = DisplayBlanksAsValues.Gap };
            ShowDataLabelsOverMaximum showDataLabelsOverMaximum1 = new ShowDataLabelsOverMaximum() { Val = false };

            chart1.Append(autoTitleDeleted1);
            chart1.Append(plotArea1);
            chart1.Append(legend1);
            chart1.Append(plotVisibleOnly1);
            chart1.Append(displayBlanksAs1);
            chart1.Append(showDataLabelsOverMaximum1);
            return chart1;
        }
    }
}