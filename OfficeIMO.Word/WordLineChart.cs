using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    public class WordLineChart : WordChart {

        public static WordLineChart AddLineChart(WordDocument wordDocument, WordParagraph paragraph) {
            throw new System.NotImplementedException();
        }

        // Creates an RoundedCorners instance and adds its children.
        public RoundedCorners GenerateRoundedCorners() {
            RoundedCorners roundedCorners1 = new RoundedCorners() { Val = false };
            return roundedCorners1;
        }
        public DocumentFormat.OpenXml.Drawing.Charts.Chart GenerateChart() {
            DocumentFormat.OpenXml.Drawing.Charts.Chart chart1 = new DocumentFormat.OpenXml.Drawing.Charts.Chart();
            AutoTitleDeleted autoTitleDeleted1 = new AutoTitleDeleted() { Val = false };

            PlotArea plotArea1 = new PlotArea();
            Layout layout1 = new Layout();

            LineChart lineChart1 = new LineChart();
            lineChart1.AddNamespaceDeclaration("c", "http://schemas.openxmlformats.org/drawingml/2006/chart");
            Grouping grouping1 = new Grouping() { Val = GroupingValues.Standard };

            DataLabels dataLabels1 = new DataLabels();
            dataLabels1.AddNamespaceDeclaration("c", "http://schemas.openxmlformats.org/drawingml/2006/chart");
            ShowLegendKey showLegendKey1 = new ShowLegendKey() { Val = false };
            ShowValue showValue1 = new ShowValue() { Val = false };
            ShowCategoryName showCategoryName1 = new ShowCategoryName() { Val = false };
            ShowSeriesName showSeriesName1 = new ShowSeriesName() { Val = false };
            ShowPercent showPercent1 = new ShowPercent() { Val = false };
            ShowBubbleSize showBubbleSize1 = new ShowBubbleSize() { Val = false };
            ShowLeaderLines showLeaderLines1 = new ShowLeaderLines() { Val = true };

            dataLabels1.Append(showLegendKey1);
            dataLabels1.Append(showValue1);
            dataLabels1.Append(showCategoryName1);
            dataLabels1.Append(showSeriesName1);
            dataLabels1.Append(showPercent1);
            dataLabels1.Append(showBubbleSize1);
            dataLabels1.Append(showLeaderLines1);

            AxisId axisId1 = new AxisId() { Val = (UInt32Value)148921728U };
            axisId1.AddNamespaceDeclaration("c", "http://schemas.openxmlformats.org/drawingml/2006/chart");

            AxisId axisId2 = new AxisId() { Val = (UInt32Value)154227840U };
            axisId2.AddNamespaceDeclaration("c", "http://schemas.openxmlformats.org/drawingml/2006/chart");

            LineChartSeries lineChartSeries1 = new LineChartSeries();
            DocumentFormat.OpenXml.Drawing.Charts.Index index1 = new DocumentFormat.OpenXml.Drawing.Charts.Index() { Val = (UInt32Value)1U };
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

            DocumentFormat.OpenXml.Drawing.Outline outline1 = new DocumentFormat.OpenXml.Drawing.Outline();

            DocumentFormat.OpenXml.Drawing.SolidFill solidFill1 = new DocumentFormat.OpenXml.Drawing.SolidFill();
            DocumentFormat.OpenXml.Drawing.RgbColorModelHex rgbColorModelHex1 = new DocumentFormat.OpenXml.Drawing.RgbColorModelHex() { Val = "FFFF00" };

            solidFill1.Append(rgbColorModelHex1);

            outline1.Append(solidFill1);

            chartShapeProperties1.Append(outline1);
            OpenXmlUnknownElement openXmlUnknownElement1 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<c:invertIfNegative xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\">0</c:invertIfNegative>");

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

            lineChartSeries1.Append(index1);
            lineChartSeries1.Append(order1);
            lineChartSeries1.Append(seriesText1);
            lineChartSeries1.Append(chartShapeProperties1);
            lineChartSeries1.Append(openXmlUnknownElement1);
            lineChartSeries1.Append(categoryAxisData1);
            lineChartSeries1.Append(values1);

            LineChartSeries lineChartSeries2 = new LineChartSeries();
            DocumentFormat.OpenXml.Drawing.Charts.Index index2 = new DocumentFormat.OpenXml.Drawing.Charts.Index() { Val = (UInt32Value)2U };
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

            DocumentFormat.OpenXml.Drawing.Outline outline2 = new DocumentFormat.OpenXml.Drawing.Outline();

            DocumentFormat.OpenXml.Drawing.SolidFill solidFill2 = new DocumentFormat.OpenXml.Drawing.SolidFill();
            DocumentFormat.OpenXml.Drawing.RgbColorModelHex rgbColorModelHex2 = new DocumentFormat.OpenXml.Drawing.RgbColorModelHex() { Val = "0000FF" };

            solidFill2.Append(rgbColorModelHex2);

            outline2.Append(solidFill2);

            chartShapeProperties2.Append(outline2);
            OpenXmlUnknownElement openXmlUnknownElement2 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<c:invertIfNegative xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\">0</c:invertIfNegative>");

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

            lineChartSeries2.Append(index2);
            lineChartSeries2.Append(order2);
            lineChartSeries2.Append(seriesText2);
            lineChartSeries2.Append(chartShapeProperties2);
            lineChartSeries2.Append(openXmlUnknownElement2);
            lineChartSeries2.Append(categoryAxisData2);
            lineChartSeries2.Append(values2);

            LineChartSeries lineChartSeries3 = new LineChartSeries();
            DocumentFormat.OpenXml.Drawing.Charts.Index index3 = new DocumentFormat.OpenXml.Drawing.Charts.Index() { Val = (UInt32Value)3U };
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

            DocumentFormat.OpenXml.Drawing.Outline outline3 = new DocumentFormat.OpenXml.Drawing.Outline();

            DocumentFormat.OpenXml.Drawing.SolidFill solidFill3 = new DocumentFormat.OpenXml.Drawing.SolidFill();
            DocumentFormat.OpenXml.Drawing.RgbColorModelHex rgbColorModelHex3 = new DocumentFormat.OpenXml.Drawing.RgbColorModelHex() { Val = "FF0000" };

            solidFill3.Append(rgbColorModelHex3);

            outline3.Append(solidFill3);

            chartShapeProperties3.Append(outline3);
            OpenXmlUnknownElement openXmlUnknownElement3 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<c:invertIfNegative xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\">0</c:invertIfNegative>");

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

            lineChartSeries3.Append(index3);
            lineChartSeries3.Append(order3);
            lineChartSeries3.Append(seriesText3);
            lineChartSeries3.Append(chartShapeProperties3);
            lineChartSeries3.Append(openXmlUnknownElement3);
            lineChartSeries3.Append(categoryAxisData3);
            lineChartSeries3.Append(values3);

            lineChart1.Append(grouping1);
            lineChart1.Append(dataLabels1);
            lineChart1.Append(axisId1);
            lineChart1.Append(axisId2);
            lineChart1.Append(lineChartSeries1);
            lineChart1.Append(lineChartSeries2);
            lineChart1.Append(lineChartSeries3);

            CategoryAxis categoryAxis1 = new CategoryAxis();
            categoryAxis1.AddNamespaceDeclaration("c", "http://schemas.openxmlformats.org/drawingml/2006/chart");
            AxisId axisId3 = new AxisId() { Val = (UInt32Value)148921728U };

            Scaling scaling1 = new Scaling();
            Orientation orientation1 = new Orientation() { Val = OrientationValues.MinMax };

            scaling1.Append(orientation1);
            Delete delete1 = new Delete() { Val = false };
            AxisPosition axisPosition1 = new AxisPosition() { Val = AxisPositionValues.Bottom };
            MajorTickMark majorTickMark1 = new MajorTickMark() { Val = TickMarkValues.Outside };
            MinorTickMark minorTickMark1 = new MinorTickMark() { Val = TickMarkValues.None };
            TickLabelPosition tickLabelPosition1 = new TickLabelPosition() { Val = TickLabelPositionValues.NextTo };
            CrossingAxis crossingAxis1 = new CrossingAxis() { Val = (UInt32Value)154227840U };
            Crosses crosses1 = new Crosses() { Val = CrossesValues.AutoZero };
            AutoLabeled autoLabeled1 = new AutoLabeled() { Val = true };
            LabelAlignment labelAlignment1 = new LabelAlignment() { Val = LabelAlignmentValues.Center };
            LabelOffset labelOffset1 = new LabelOffset() { Val = (UInt16Value)100U };
            NoMultiLevelLabels noMultiLevelLabels1 = new NoMultiLevelLabels() { Val = false };

            categoryAxis1.Append(axisId3);
            categoryAxis1.Append(scaling1);
            categoryAxis1.Append(delete1);
            categoryAxis1.Append(axisPosition1);
            categoryAxis1.Append(majorTickMark1);
            categoryAxis1.Append(minorTickMark1);
            categoryAxis1.Append(tickLabelPosition1);
            categoryAxis1.Append(crossingAxis1);
            categoryAxis1.Append(crosses1);
            categoryAxis1.Append(autoLabeled1);
            categoryAxis1.Append(labelAlignment1);
            categoryAxis1.Append(labelOffset1);
            categoryAxis1.Append(noMultiLevelLabels1);

            ValueAxis valueAxis1 = new ValueAxis();
            valueAxis1.AddNamespaceDeclaration("c", "http://schemas.openxmlformats.org/drawingml/2006/chart");
            AxisId axisId4 = new AxisId() { Val = (UInt32Value)154227840U };

            Scaling scaling2 = new Scaling();
            Orientation orientation2 = new Orientation() { Val = OrientationValues.MinMax };

            scaling2.Append(orientation2);
            Delete delete2 = new Delete() { Val = false };
            AxisPosition axisPosition2 = new AxisPosition() { Val = AxisPositionValues.Left };
            DocumentFormat.OpenXml.Drawing.Charts.NumberingFormat numberingFormat1 = new DocumentFormat.OpenXml.Drawing.Charts.NumberingFormat() { FormatCode = "General", SourceLinked = false };
            MajorGridlines majorGridlines1 = new MajorGridlines();
            MajorTickMark majorTickMark2 = new MajorTickMark() { Val = TickMarkValues.Outside };
            MinorTickMark minorTickMark2 = new MinorTickMark() { Val = TickMarkValues.None };
            TickLabelPosition tickLabelPosition2 = new TickLabelPosition() { Val = TickLabelPositionValues.NextTo };
            CrossingAxis crossingAxis2 = new CrossingAxis() { Val = (UInt32Value)148921728U };
            Crosses crosses2 = new Crosses() { Val = CrossesValues.AutoZero };
            CrossBetween crossBetween1 = new CrossBetween() { Val = CrossBetweenValues.Between };

            valueAxis1.Append(axisId4);
            valueAxis1.Append(scaling2);
            valueAxis1.Append(delete2);
            valueAxis1.Append(axisPosition2);
            valueAxis1.Append(numberingFormat1);
            valueAxis1.Append(majorGridlines1);
            valueAxis1.Append(majorTickMark2);
            valueAxis1.Append(minorTickMark2);
            valueAxis1.Append(tickLabelPosition2);
            valueAxis1.Append(crossingAxis2);
            valueAxis1.Append(crosses2);
            valueAxis1.Append(crossBetween1);

            plotArea1.Append(layout1);
            plotArea1.Append(lineChart1);
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

        public DocumentFormat.OpenXml.Wordprocessing.Run GenerateRun() {
            DocumentFormat.OpenXml.Wordprocessing.Run run1 = new DocumentFormat.OpenXml.Wordprocessing.Run();
            run1.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            Drawing drawing1 = new Drawing();

            DocumentFormat.OpenXml.Drawing.Wordprocessing.Inline inline1 = new DocumentFormat.OpenXml.Drawing.Wordprocessing.Inline();
            inline1.AddNamespaceDeclaration("wp", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing");
            DocumentFormat.OpenXml.Drawing.Wordprocessing.Extent extent1 = new DocumentFormat.OpenXml.Drawing.Wordprocessing.Extent() { Cx = 5486400L, Cy = 3200400L };
            DocumentFormat.OpenXml.Drawing.Wordprocessing.EffectExtent effectExtent1 = new DocumentFormat.OpenXml.Drawing.Wordprocessing.EffectExtent() { LeftEdge = 0L, TopEdge = 0L, RightEdge = 19050L, BottomEdge = 19050L };
            DocumentFormat.OpenXml.Drawing.Wordprocessing.DocProperties docProperties1 = new DocumentFormat.OpenXml.Drawing.Wordprocessing.DocProperties() { Id = (UInt32Value)2U, Name = "chart" };

            DocumentFormat.OpenXml.Drawing.Graphic graphic1 = new DocumentFormat.OpenXml.Drawing.Graphic();
            graphic1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

            DocumentFormat.OpenXml.Drawing.GraphicData graphicData1 = new DocumentFormat.OpenXml.Drawing.GraphicData() { Uri = "http://schemas.openxmlformats.org/drawingml/2006/chart" };

            DocumentFormat.OpenXml.Drawing.Charts.ChartReference chartReference1 = new DocumentFormat.OpenXml.Drawing.Charts.ChartReference() { Id = "rId1" };
            chartReference1.AddNamespaceDeclaration("p6", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            chartReference1.AddNamespaceDeclaration("c", "http://schemas.openxmlformats.org/drawingml/2006/chart");

            graphicData1.Append(chartReference1);

            graphic1.Append(graphicData1);

            inline1.Append(extent1);
            inline1.Append(effectExtent1);
            inline1.Append(docProperties1);
            inline1.Append(graphic1);

            drawing1.Append(inline1);

            run1.Append(drawing1);
            return run1;
        }

    }
}
