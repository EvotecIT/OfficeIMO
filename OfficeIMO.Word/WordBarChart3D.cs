using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    public class WordBarChart3D : WordChart {

        public static WordBarChart3D AddBarChart3D(WordDocument wordDocument, WordParagraph paragraph) {
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

            Bar3DChart bar3DChart1 = new Bar3DChart();
            bar3DChart1.AddNamespaceDeclaration("c", "http://schemas.openxmlformats.org/drawingml/2006/chart");
            BarDirection barDirection1 = new BarDirection() { Val = BarDirectionValues.Column };
            BarGrouping barGrouping1 = new BarGrouping() { Val = BarGroupingValues.Clustered };
            GapWidth gapWidth1 = new GapWidth() { Val = (UInt16Value)150U };

            AxisId axisId1 = new AxisId() { Val = (UInt32Value)148921728U };
            axisId1.AddNamespaceDeclaration("c", "http://schemas.openxmlformats.org/drawingml/2006/chart");

            AxisId axisId2 = new AxisId() { Val = (UInt32Value)154227840U };
            axisId2.AddNamespaceDeclaration("c", "http://schemas.openxmlformats.org/drawingml/2006/chart");

            OpenXmlUnknownElement openXmlUnknownElement1 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<c:overlap val=\"0\" xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\" />");

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

            bar3DChart1.Append(barDirection1);
            bar3DChart1.Append(barGrouping1);
            bar3DChart1.Append(gapWidth1);
            bar3DChart1.Append(axisId1);
            bar3DChart1.Append(axisId2);
            bar3DChart1.Append(openXmlUnknownElement1);
            bar3DChart1.Append(dataLabels1);
            bar3DChart1.Append(barChartSeries1);

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
            plotArea1.Append(bar3DChart1);
            plotArea1.Append(categoryAxis1);
            plotArea1.Append(valueAxis1);
            PlotVisibleOnly plotVisibleOnly1 = new PlotVisibleOnly() { Val = true };
            DisplayBlanksAs displayBlanksAs1 = new DisplayBlanksAs() { Val = DisplayBlanksAsValues.Gap };
            ShowDataLabelsOverMaximum showDataLabelsOverMaximum1 = new ShowDataLabelsOverMaximum() { Val = false };

            chart1.Append(autoTitleDeleted1);
            chart1.Append(plotArea1);
            chart1.Append(plotVisibleOnly1);
            chart1.Append(displayBlanksAs1);
            chart1.Append(showDataLabelsOverMaximum1);
            return chart1;
        }

        public Run GenerateRun() {
            Run run1 = new Run();
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


        public WordBarChart3D(WordDocument document, Paragraph paragraph, Drawing drawing) : base(document, paragraph, drawing) {
        }
    }
}
