using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Core.Internal.OpenXml;
using static OfficeIMO.OpenXml.Internal.OfficeChartXml;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.OpenXml.Internal {
    internal static partial class OfficeOpenXmlChartSeriesReader {
        internal static bool HasUnsupportedBubblePresentation(IOfficeChartPart part, OfficeXmlNode chart, OfficeXmlNode plotArea,
            OfficeXmlNode bubble, int maximumPointOverrides) =>
            IsEnabledFlag(bubble.Child(NsC, "varyColors")) ||
            IsEnabledFlag(bubble.Child(NsC, "bubble3D")) ||
            HasUnsupportedBubbleAxes(plotArea, bubble) || HasUnsupportedBubbleLegend(chart) ||
            HasUnsupportedBubbleAreaLayout(part, plotArea) || HasEnabledBubbleDataLabels(bubble) ||
            bubble.Children(NsC, "ser").ToList().Any(series => IsEnabledFlag(series.Child(NsC, "bubble3D")) ||
                OfficeOpenXmlChartPointStyles.GetBoundedPoints(series, maximumPointOverrides)
                    .Any(point => IsEnabledFlag(point.Child(NsC, "bubble3D"))));

        // c:bubble3D and c:varyColors: present with a val other than false.
        private static bool IsEnabledFlag(OfficeXmlNode flag) =>
            !flag.IsNone && flag.BoolVal() != false;

        private static bool HasUnsupportedBubbleAxes(
            OfficeXmlNode plotArea, OfficeXmlNode chart) {
            if (!TryGetReferencedBubbleAxes(
                    plotArea, chart, out OfficeXmlNode horizontalAxis,
                    out OfficeXmlNode verticalAxis)) {
                return true;
            }
            if (horizontalAxis.ChildVal(NsC, "axPos") != "b" ||
                verticalAxis.ChildVal(NsC, "axPos") != "l") {
                return true;
            }
            if (!HasSupportedDefaultBubbleGridlines(
                    horizontalAxis, verticalAxis)) {
                return true;
            }
            return new[] { horizontalAxis, verticalAxis }.Any(axis =>
                HasUnsupportedBubbleAxisPresentation(axis) ||
                (axis.Child(NsC, "delete") is { IsNone: false } delete &&
                  delete.BoolVal() != false) ||
                 !axis.Child(NsC, "majorUnit").IsNone ||
                 !axis.Child(NsC, "minorUnit").IsNone ||
                 !axis.Child(NsC, "dispUnits").IsNone ||
                 !axis.Child(NsC, "crossesAt").IsNone ||
                 OfficeOpenXmlChartSecondaryAxis.IsWorkbookLinked(axis) ||
                 (axis.Child(NsC, "tickLblPos") is { IsNone: false } tickLabelPosition &&
                  tickLabelPosition.Val() != "nextTo") ||
                 (axis.Child(NsC, "crosses") is { IsNone: false } crosses &&
                  crosses.Val() != "autoZero") ||
                 (axis.Child(NsC, "scaling") is { IsNone: false } scaling &&
                  (!scaling.Child(NsC, "logBase").IsNone ||
                   !scaling.Child(NsC, "min").IsNone ||
                   !scaling.Child(NsC, "max").IsNone ||
                   scaling.ChildVal(NsC, "orientation") == "maxMin")));
        }

        private static bool HasUnsupportedBubbleAxisPresentation(
            OfficeXmlNode axis) =>
            HasUnsupportedBubbleTitle(axis.Child(NsC, "title")) ||
            HasUnsupportedBubbleTextStyle(axis) ||
            HasUnsupportedBubbleShapeProperties(axis);

        private static bool HasSupportedDefaultBubbleGridlines(
            OfficeXmlNode horizontalAxis, OfficeXmlNode verticalAxis) {
            if (!horizontalAxis.Child(NsC, "majorGridlines").IsNone ||
                !horizontalAxis.Child(NsC, "minorGridlines").IsNone ||
                !verticalAxis.Child(NsC, "minorGridlines").IsNone) {
                return false;
            }

            OfficeXmlNode gridlines =
                verticalAxis.Child(NsC, "majorGridlines");
            OfficeXmlNode properties =
                gridlines.Child(NsC, "spPr");
            if (gridlines.IsNone || properties.IsNone) return false;
            return !HasUnsupportedBubbleShapeProperties(gridlines);
        }

        internal static bool TryGetReferencedBubbleAxes(
            OfficeXmlNode plotArea, OfficeXmlNode chart,
            out OfficeXmlNode horizontalAxis, out OfficeXmlNode verticalAxis) {
            horizontalAxis = OfficeXmlNode.None;
            verticalAxis = OfficeXmlNode.None;
            List<OfficeXmlNode> references =
                chart.Children(NsC, "axId").ToList();
            if (references.Count != 2 ||
                references.Any(axis => axis.UIntVal() == null)) {
                return false;
            }
            uint horizontalId = references[0].UIntVal()!.Value;
            uint verticalId = references[1].UIntVal()!.Value;
            if (horizontalId == verticalId) return false;
            OfficeXmlNode horizontal = plotArea.Children(NsC, "valAx").ToList()
                .FirstOrDefault(axis =>
                    axis.Child(NsC, "axId").UIntVal() == horizontalId);
            OfficeXmlNode vertical = plotArea.Children(NsC, "valAx").ToList()
                .FirstOrDefault(axis =>
                    axis.Child(NsC, "axId").UIntVal() == verticalId);
            if (horizontal.IsNone || vertical.IsNone) return false;
            horizontalAxis = horizontal;
            verticalAxis = vertical;
            return true;
        }

        private static bool HasUnsupportedBubbleLegend(OfficeXmlNode chart) {
            OfficeXmlNode legend = chart.Child(NsC, "legend");
            return !legend.IsNone &&
                (legend.ChildVal(NsC, "legendPos") == "tr" ||
                 !legend.Child(NsC, "layout")
                     .Child(NsC, "manualLayout").IsNone ||
                 HasUnsupportedBubbleTextStyle(legend) ||
                 HasUnsupportedBubbleShapeProperties(legend));
        }

        private static bool HasUnsupportedBubbleAreaLayout(
            IOfficeChartPart chartPart, OfficeXmlNode plotArea) =>
            HasUnsupportedBubbleTitle(
                chartPart.ChartSpace.Child(NsC, "chart")
                    .Child(NsC, "title")) ||
            !plotArea.Child(NsC, "layout")
                .Child(NsC, "manualLayout").IsNone ||
            !IsFramelessChartArea(chartPart.ChartSpace.Child(NsC, "spPr")) ||
            plotArea.Child(NsC, "spPr").HasChildElements();

        // No chart-area shape, or one with no fill and no outline, draws nothing.
        private static bool IsFramelessChartArea(OfficeXmlNode shape) =>
            shape.IsNone || shape.Children().ToList().All(child => child.IsDrawing("noFill") ||
                child.IsDrawing("ln") && !child.HasAttributes && child.Children().ToList().All(item => item.IsDrawing("noFill")));

        private static bool HasUnsupportedBubbleTitle(OfficeXmlNode title) =>
            !title.IsNone &&
            (!title.Child(NsC, "layout")
                 .Child(NsC, "manualLayout").IsNone ||
             HasUnsupportedBubbleTextStyle(title) ||
             HasUnsupportedBubbleShapeProperties(title));

        private static bool HasUnsupportedBubbleTextStyle(OfficeXmlNode parent) {
            OfficeXmlNode chart = parent.IsChart("chart") ? parent : parent.Ancestor(NsC, "chart");
            if (chart.IsNone) return true;
            try {
                OfficeOpenXmlChartTextReader.ReadNativeText(chart, parent, scheme: OfficeXmlNode.None, allowMultiline: true);
                return false;
            } catch (NotSupportedException) {
                return true;
            }
        }

        private static bool HasUnsupportedBubbleShapeProperties(OfficeXmlNode parent) {
            try {
                ReadSurface(parent.Child(NsC, "spPr"), scheme: OfficeXmlNode.None, colorMap: OfficeXmlNode.None);
                return false;
            } catch (NotSupportedException) {
                return true;
            }
        }

        private static bool HasEnabledBubbleDataLabels(OfficeXmlNode chart) =>
            chart.Descendants(NsC, "showLegendKey").Any(item => item.BoolVal() != false) ||
            chart.Descendants(NsC, "showVal").Any(item => item.BoolVal() != false) ||
            chart.Descendants(NsC, "showCatName").Any(item => item.BoolVal() != false) ||
            chart.Descendants(NsC, "showSerName").Any(item => item.BoolVal() != false) ||
            chart.Descendants(NsC, "showPercent").Any(item => item.BoolVal() != false) ||
            chart.Descendants(NsC, "showBubbleSize").Any(item => item.BoolVal() != false) ||
            chart.Descendants(NsC, "dLbl").Any(label => {
                OfficeXmlNode delete = label.Child(NsC, "delete");
                return !label.Child(NsC, "tx").IsNone &&
                    (delete.IsNone || delete.BoolVal() == false);
            });


        internal static bool HasUnsupportedSharedAxisNumberFormat(OfficeXmlNode axis) =>
            OfficeOpenXmlChartSecondaryAxis.HasUnsupportedSharedAxisNumberFormat(axis);

        // SDK callers
        internal static bool HasUnsupportedBubblePresentation(ChartPart part, C.Chart chart, C.PlotArea plotArea,
            C.BubbleChart bubble, int maximumPointOverrides) =>
            HasUnsupportedBubblePresentation(new OfficeSdkChartPart(part), OfficeXmlNode.FromSdk(chart), OfficeXmlNode.FromSdk(plotArea),
                OfficeXmlNode.FromSdk(bubble), maximumPointOverrides);

        internal static bool TryGetReferencedBubbleAxes(
            C.PlotArea plotArea, C.BubbleChart chart,
            out C.ValueAxis horizontalAxis, out C.ValueAxis verticalAxis) {
            bool found = TryGetReferencedBubbleAxes(OfficeXmlNode.FromSdk(plotArea), OfficeXmlNode.FromSdk(chart),
                out OfficeXmlNode horizontal, out OfficeXmlNode vertical);
            horizontalAxis = (C.ValueAxis)horizontal.SdkElement!;
            verticalAxis = (C.ValueAxis)vertical.SdkElement!;
            return found;
        }
    }
}
