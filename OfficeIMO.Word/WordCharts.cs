using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Office2016.Drawing.ChartDrawing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using AxisId = DocumentFormat.OpenXml.Drawing.Charts.AxisId;
using Chart = DocumentFormat.OpenXml.Drawing.Charts.Chart;
using ChartSpace = DocumentFormat.OpenXml.Drawing.Charts.ChartSpace;
using DataLabels = DocumentFormat.OpenXml.Drawing.Charts.DataLabels;
using Legend = DocumentFormat.OpenXml.Drawing.Charts.Legend;
using PlotArea = DocumentFormat.OpenXml.Drawing.Charts.PlotArea;

namespace OfficeIMO.Word {
    public partial class WordChart {
        protected static WordDocument _document;
        protected static WordParagraph _paragraph;
        protected static ChartPart _chartPart;
        private string _id {
            get {
                return _document._wordprocessingDocument.MainDocumentPart.GetIdOfPart(_chartPart);
            }
        }

        internal static CategoryAxis AddCategoryAxis() {
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

            return categoryAxis1;
        }

        internal static Legend AddLegend() {
            Legend legend1 = new Legend();
            LegendPosition legendPosition1 = new LegendPosition() { Val = LegendPositionValues.Left };
            Overlay overlay1 = new Overlay() { Val = false };

            legend1.Append(legendPosition1);
            legend1.Append(overlay1);
            return legend1;
        }

        internal static ValueAxis AddValueAxis() {
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

            return valueAxis1;
        }

        internal static DataLabels AddDataLabel() {
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
            return dataLabels1;
        }

        internal static WordParagraph InsertChart(WordDocument wordDocument, WordParagraph paragraph, Chart chart, bool roundedCorners) {
            ChartPart part = CreateChartPart(wordDocument, roundedCorners);
            _chartPart = part;
            var id = _document._wordprocessingDocument.MainDocumentPart.GetIdOfPart(_chartPart);

            Drawing chartDrawing = CreateChartDrawing(id);

            var run = new Run();
            run.Append(chartDrawing);
            paragraph._paragraph.Append(run);
            _chartPart.ChartSpace.Append(chart);
            return paragraph;
        }

        internal static ChartPart CreateChartPart(WordDocument document, bool roundedCorners) {
            ChartPart part = document._wordprocessingDocument.MainDocumentPart.AddNewPart<ChartPart>(); //("rId1");

            ChartSpace chartSpace1 = new ChartSpace();
            chartSpace1.AddNamespaceDeclaration("c", "http://schemas.openxmlformats.org/drawingml/2006/chart");
            chartSpace1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
            chartSpace1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");

            part.ChartSpace = chartSpace1;
            part.ChartSpace.Append(new RoundedCorners() { Val = roundedCorners });
            return part;
        }

        internal static Chart GenerateChart() {
            Chart chart1 = new Chart();
            AutoTitleDeleted autoTitleDeleted1 = new AutoTitleDeleted() { Val = false };
            PlotArea plotArea1 = new PlotArea();
            Layout layout1 = new Layout();
            plotArea1.Append(layout1);
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
            return chart1;
        }

        internal static Drawing CreateChartDrawing(string id) {
            Drawing drawing1 = new Drawing();

            DocumentFormat.OpenXml.Drawing.Wordprocessing.Inline inline1 = new DocumentFormat.OpenXml.Drawing.Wordprocessing.Inline();
            inline1.AddNamespaceDeclaration("wp", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing");
            DocumentFormat.OpenXml.Drawing.Wordprocessing.Extent extent1 = new DocumentFormat.OpenXml.Drawing.Wordprocessing.Extent() { Cx = 4445000L, Cy = 6985000L };
            DocumentFormat.OpenXml.Drawing.Wordprocessing.EffectExtent effectExtent1 = new DocumentFormat.OpenXml.Drawing.Wordprocessing.EffectExtent() { LeftEdge = 0L, TopEdge = 0L, RightEdge = 19050L, BottomEdge = 19050L };
            DocumentFormat.OpenXml.Drawing.Wordprocessing.DocProperties docProperties1 = new DocumentFormat.OpenXml.Drawing.Wordprocessing.DocProperties() { Id = (UInt32Value)2U, Name = "chart" };

            DocumentFormat.OpenXml.Drawing.Graphic graphic1 = new DocumentFormat.OpenXml.Drawing.Graphic();
            graphic1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

            DocumentFormat.OpenXml.Drawing.GraphicData graphicData1 = new DocumentFormat.OpenXml.Drawing.GraphicData() { Uri = "http://schemas.openxmlformats.org/drawingml/2006/chart" };

            DocumentFormat.OpenXml.Drawing.Charts.ChartReference chartReference1 = new DocumentFormat.OpenXml.Drawing.Charts.ChartReference() { Id = id };
            chartReference1.AddNamespaceDeclaration("p6", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            chartReference1.AddNamespaceDeclaration("c", "http://schemas.openxmlformats.org/drawingml/2006/chart");

            graphicData1.Append(chartReference1);

            graphic1.Append(graphicData1);

            inline1.Append(extent1);
            inline1.Append(effectExtent1);
            inline1.Append(docProperties1);
            inline1.Append(graphic1);

            drawing1.Append(inline1);
            return drawing1;
        }

    }
}
