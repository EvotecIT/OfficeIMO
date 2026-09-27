using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Drawing;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.OpenXml.Internal {
    internal static partial class OfficeOpenXmlChartSeriesReader {
        internal static bool HasUnsupportedBubblePresentation(ChartPart part, C.Chart chart, C.PlotArea plotArea,
            C.BubbleChart bubble, int maximumPoints) =>
            IsVaryColorsEnabled(bubble.GetFirstChild<C.VaryColors>()) ||
            IsBubble3DEnabled(bubble.GetFirstChild<C.Bubble3D>()) ||
            HasUnsupportedBubbleSourceVisibility(part, chart) ||
            HasUnsupportedBubbleAxes(plotArea, bubble) || HasUnsupportedBubbleLegend(chart) ||
            HasUnsupportedBubbleAreaLayout(part, plotArea) || HasEnabledBubbleDataLabels(bubble) ||
            bubble.Elements<C.BubbleChartSeries>().Any(series => IsBubble3DEnabled(series.GetFirstChild<C.Bubble3D>()) ||
                OfficeOpenXmlChartCacheReader.GetBoundedCachedPoints(series.Elements<C.DataPoint>(), maximumPoints)
                    .Any(point => IsBubble3DEnabled(point.GetFirstChild<C.Bubble3D>())));

        private static bool IsBubble3DEnabled(C.Bubble3D? bubble3D) =>
            bubble3D != null && bubble3D.Val?.Value != false;

        private static bool IsVaryColorsEnabled(C.VaryColors? varyColors) =>
            varyColors != null && varyColors.Val?.Value != false;

        private static bool HasUnsupportedBubbleAxes(
            C.PlotArea plotArea, C.BubbleChart chart) {
            if (!TryGetReferencedBubbleAxes(
                    plotArea, chart, out C.ValueAxis horizontalAxis,
                    out C.ValueAxis verticalAxis)) {
                return true;
            }
            if (horizontalAxis.AxisPosition?.Val?.Value !=
                    C.AxisPositionValues.Bottom ||
                verticalAxis.AxisPosition?.Val?.Value !=
                    C.AxisPositionValues.Left) {
                return true;
            }
            if (!HasSupportedDefaultBubbleGridlines(
                    horizontalAxis, verticalAxis)) {
                return true;
            }
            return new[] { horizontalAxis, verticalAxis }.Any(axis =>
                HasUnsupportedBubbleAxisPresentation(axis) ||
                (axis.GetFirstChild<C.Delete>() is C.Delete delete &&
                  delete.Val?.Value != false) ||
                 axis.GetFirstChild<C.MajorUnit>() != null ||
                 axis.GetFirstChild<C.MinorUnit>() != null ||
                 axis.GetFirstChild<C.DisplayUnits>() != null ||
                 axis.GetFirstChild<C.CrossesAt>() != null ||
                 HasUnsupportedSharedAxisNumberFormat(axis) ||
                 (axis.GetFirstChild<C.TickLabelPosition>() is
                      C.TickLabelPosition tickLabelPosition &&
                  tickLabelPosition.Val?.Value !=
                      C.TickLabelPositionValues.NextTo) ||
                 (axis.GetFirstChild<C.Crosses>() is C.Crosses crosses &&
                  crosses.Val?.Value != C.CrossesValues.AutoZero) ||
                 (axis.GetFirstChild<C.Scaling>() is C.Scaling scaling &&
                  (scaling.GetFirstChild<C.LogBase>() != null ||
                   scaling.GetFirstChild<C.MinAxisValue>() != null ||
                   scaling.GetFirstChild<C.MaxAxisValue>() != null ||
                   scaling.GetFirstChild<C.Orientation>()?.Val?.Value ==
                      C.OrientationValues.MaxMin)));
        }

        private static bool HasUnsupportedBubbleAxisPresentation(
            C.ValueAxis axis) =>
            HasUnsupportedBubbleTitle(axis.GetFirstChild<C.Title>()) ||
            HasUnsupportedBubbleTextStyle(axis) ||
            HasUnsupportedBubbleShapeProperties(axis);

        private static bool HasSupportedDefaultBubbleGridlines(
            C.ValueAxis horizontalAxis, C.ValueAxis verticalAxis) {
            if (horizontalAxis.GetFirstChild<C.MajorGridlines>() != null ||
                horizontalAxis.GetFirstChild<C.MinorGridlines>() != null ||
                verticalAxis.GetFirstChild<C.MinorGridlines>() != null) {
                return false;
            }

            C.MajorGridlines? gridlines =
                verticalAxis.GetFirstChild<C.MajorGridlines>();
            C.ChartShapeProperties? properties =
                gridlines?.GetFirstChild<C.ChartShapeProperties>();
            A.Outline? outline = properties?.GetFirstChild<A.Outline>();
            if (gridlines == null || properties == null || outline == null ||
                gridlines.ChildElements.Count != 1 ||
                properties.ChildElements.Count != 1 ||
                outline.ChildElements.Count != 1 ||
                outline.Width?.Value !=
                    6350L) {
                return false;
            }

            OfficeColor? color = OfficeOpenXmlThemeColorResolver.ResolveColor(
                outline.GetFirstChild<A.SolidFill>(), colorScheme: null);
            return color == OfficeChartStyle.Default.GridLineColor;
        }

        internal static bool TryGetReferencedBubbleAxes(
            C.PlotArea plotArea, C.BubbleChart chart,
            out C.ValueAxis horizontalAxis, out C.ValueAxis verticalAxis) {
            horizontalAxis = null!;
            verticalAxis = null!;
            List<C.AxisId> references =
                chart.Elements<C.AxisId>().ToList();
            if (references.Count != 2 ||
                references.Any(axis => axis.Val?.Value == null)) {
                return false;
            }
            uint horizontalId = references[0].Val!.Value;
            uint verticalId = references[1].Val!.Value;
            if (horizontalId == verticalId) return false;
            C.ValueAxis? horizontal = plotArea.Elements<C.ValueAxis>()
                .FirstOrDefault(axis =>
                    axis.AxisId?.Val?.Value == horizontalId);
            C.ValueAxis? vertical = plotArea.Elements<C.ValueAxis>()
                .FirstOrDefault(axis =>
                    axis.AxisId?.Val?.Value == verticalId);
            if (horizontal == null || vertical == null) return false;
            horizontalAxis = horizontal;
            verticalAxis = vertical;
            return true;
        }

        private static bool HasUnsupportedBubbleLegend(C.Chart chart) {
            C.Legend? legend = chart.GetFirstChild<C.Legend>();
            return legend != null &&
                (legend.GetFirstChild<C.LegendPosition>()?.Val?.Value ==
                     C.LegendPositionValues.TopRight ||
                 legend.GetFirstChild<C.Layout>()?
                     .GetFirstChild<C.ManualLayout>() != null ||
                 HasUnsupportedBubbleTextStyle(legend) ||
                 HasUnsupportedBubbleShapeProperties(legend));
        }

        private static bool HasUnsupportedBubbleAreaLayout(
            ChartPart chartPart, C.PlotArea plotArea) =>
            HasUnsupportedBubbleTitle(
                chartPart.ChartSpace?.GetFirstChild<C.Chart>()?
                    .GetFirstChild<C.Title>()) ||
            plotArea.GetFirstChild<C.Layout>()?
                .GetFirstChild<C.ManualLayout>() != null ||
            chartPart.ChartSpace?.GetFirstChild<C.ShapeProperties>()?
                .ChildElements.Count > 0 ||
            plotArea.GetFirstChild<C.ShapeProperties>()?
                .ChildElements.Count > 0;

        private static bool HasUnsupportedBubbleTitle(C.Title? title) =>
            title != null &&
            (title.GetFirstChild<C.Layout>()?
                 .GetFirstChild<C.ManualLayout>() != null ||
             HasUnsupportedBubbleTextStyle(title) ||
             HasUnsupportedBubbleShapeProperties(title));

        private static bool HasUnsupportedBubbleTextStyle(
            OpenXmlElement parent) =>
            parent.Descendants<A.RunProperties>()
                .Any(HasUnsupportedBubbleTextCharacterProperties) ||
            parent.Descendants<A.DefaultRunProperties>()
                .Any(HasUnsupportedBubbleTextCharacterProperties) ||
            parent.Descendants<A.EndParagraphRunProperties>()
                .Any(HasUnsupportedBubbleTextCharacterProperties) ||
            parent.Descendants<A.BodyProperties>()
                .Any(properties =>
                    properties.HasAttributes ||
                    properties.ChildElements.Count > 0) ||
            parent.Descendants<A.ListStyle>()
                .Any(style => style.ChildElements.Count > 0) ||
            parent.Descendants<A.ParagraphProperties>()
                .Any(properties =>
                    properties.HasAttributes ||
                    properties.ChildElements.Any(child =>
                        child is not A.DefaultRunProperties));

        private static bool HasUnsupportedBubbleTextCharacterProperties(
            A.TextCharacterPropertiesType properties) =>
            properties.ChildElements.Count > 0 ||
            properties.GetAttributes().Any(attribute =>
                !string.Equals(
                    attribute.LocalName, "lang",
                    StringComparison.Ordinal));

        private static bool HasUnsupportedBubbleShapeProperties(
            OpenXmlElement parent) {
            C.ChartShapeProperties? properties =
                parent.GetFirstChild<C.ChartShapeProperties>();
            return properties != null &&
                (properties.HasAttributes ||
                 properties.ChildElements.Count > 0);
        }

        private static bool HasEnabledBubbleDataLabels(C.BubbleChart chart) =>
            chart.Descendants<C.ShowLegendKey>().Any(item => item.Val?.Value != false) ||
            chart.Descendants<C.ShowValue>().Any(item => item.Val?.Value != false) ||
            chart.Descendants<C.ShowCategoryName>().Any(item => item.Val?.Value != false) ||
            chart.Descendants<C.ShowSeriesName>().Any(item => item.Val?.Value != false) ||
            chart.Descendants<C.ShowPercent>().Any(item => item.Val?.Value != false) ||
            chart.Descendants<C.ShowBubbleSize>().Any(item => item.Val?.Value != false) ||
            chart.Descendants<C.DataLabel>().Any(label => {
                C.Delete? delete = label.GetFirstChild<C.Delete>();
                return label.GetFirstChild<C.ChartText>() != null &&
                    (delete == null || delete.Val?.Value == false);
            });


        private static bool HasUnsupportedSharedAxisNumberFormat(
            C.ValueAxis axis) {
            string? format = axis.GetFirstChild<C.NumberingFormat>()?.FormatCode?.Value;
            if (string.IsNullOrWhiteSpace(format)) return false;
            if (string.Equals(format, "General",
                    StringComparison.OrdinalIgnoreCase)) {
                return false;
            }

            bool inQuotedLiteral = false;
            bool escaped = false;
            bool sectionHasPlaceholder = false;
            for (int index = 0; index < format!.Length; index++) {
                char value = format[index];
                if (escaped) {
                    escaped = false;
                    continue;
                }
                if (value == '\\') {
                    escaped = true;
                    continue;
                }
                if (value == '"') {
                    inQuotedLiteral = !inQuotedLiteral;
                    continue;
                }
                if (inQuotedLiteral) {
                    continue;
                }
                if (value == '0' || value == '#' || value == '?') {
                    sectionHasPlaceholder = true;
                    continue;
                }
                if (value == ';') {
                    if (!sectionHasPlaceholder) return true;
                    sectionHasPlaceholder = false;
                    continue;
                }
                if (value == '/' || value == '@' ||
                    value == '[' || value == ']') {
                    return true;
                }
                if (value != 'E' && value != 'e') continue;

                int next = index + 1;
                if (next < format.Length &&
                    (format[next] == '+' || format[next] == '-')) {
                    next++;
                }
                if (next < format.Length &&
                    (format[next] == '0' || format[next] == '#' ||
                     format[next] == '?')) {
                    return true;
                }
            }

            return inQuotedLiteral || escaped || !sectionHasPlaceholder;
        }


    }
}
