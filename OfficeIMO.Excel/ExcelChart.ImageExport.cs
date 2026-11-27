using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.Drawing;
using OfficeIMO.Excel.OpenXml;
using OfficeIMO.OpenXml.Internal;
using OfficeIMO.Excel.Utilities;
using static OfficeIMO.Excel.OpenXml.ExcelChartXml;

namespace OfficeIMO.Excel {
    public sealed partial class ExcelChart {
        private OfficeChartStyle? CreateImageExportStyle(int radialPointCount) {
            OfficeXmlNode chartSpace = ChartSpaceNode();
            if (chartSpace.IsNone) {
                return null;
            }

            OfficeXmlNode chart = chartSpace.Child(C, "chart");
            OfficeXmlNode plotArea = chart.Child(C, "plotArea");
            OfficeXmlNode chartAreaProperties = chartSpace.Child(C, "spPr");
            OfficeXmlNode plotAreaProperties = plotArea.Child(C, "spPr");
            ExcelPartRef workbookPart = ExcelPartRef.Workbook(_document);

            OfficeColor? chartFill = TryGetSolidFill(chartAreaProperties, workbookPart, out OfficeColor resolvedChartFill) ? resolvedChartFill : null;
            OfficeColor? chartLine = TryGetSolidLine(chartAreaProperties, workbookPart, out OfficeColor resolvedChartLine) ? resolvedChartLine : null;
            OfficeColor? plotFill = TryGetSolidFill(plotAreaProperties, workbookPart, out OfficeColor resolvedPlotFill) ? resolvedPlotFill : null;
            OfficeColor? plotLine = TryGetSolidLine(plotAreaProperties, workbookPart, out OfficeColor resolvedPlotLine) ? resolvedPlotLine : null;
            OfficeXmlNode categoryAxis = plotArea.IsNone ? OfficeXmlNode.None : ResolveImageExportCategoryAxis(plotArea);
            OfficeXmlNode valueAxis = plotArea.IsNone ? OfficeXmlNode.None : ResolveImageExportValueAxis(plotArea);
            double? chartLineWidth = TryGetImageExportLineWidth(chartAreaProperties, out double resolvedChartLineWidth) ? resolvedChartLineWidth : null;
            double? plotLineWidth = TryGetImageExportLineWidth(plotAreaProperties, out double resolvedPlotLineWidth) ? resolvedPlotLineWidth : null;
            OfficeStrokeDashStyle? chartLineDashStyle = TryGetLineDashStyle(chartAreaProperties, out OfficeStrokeDashStyle resolvedChartLineDashStyle) ? resolvedChartLineDashStyle : null;
            OfficeStrokeDashStyle? plotLineDashStyle = TryGetLineDashStyle(plotAreaProperties, out OfficeStrokeDashStyle resolvedPlotLineDashStyle) ? resolvedPlotLineDashStyle : null;
            OfficeColor? categoryAxisLine = TryGetImageExportAxisLineColor(categoryAxis, workbookPart, out OfficeColor resolvedCategoryAxisLine) ? resolvedCategoryAxisLine : null;
            OfficeColor? valueAxisLine = TryGetImageExportAxisLineColor(valueAxis, workbookPart, out OfficeColor resolvedValueAxisLine) ? resolvedValueAxisLine : null;
            OfficeColor? categoryGridLine = TryGetImageExportGridLineColor(categoryAxis, workbookPart, out OfficeColor resolvedCategoryGridLine) ? resolvedCategoryGridLine : null;
            OfficeColor? valueGridLine = TryGetImageExportGridLineColor(valueAxis, workbookPart, out OfficeColor resolvedValueGridLine) ? resolvedValueGridLine : null;
            OfficeColor? categoryMinorGridLine = TryGetImageExportMinorGridLineColor(categoryAxis, workbookPart, out OfficeColor resolvedCategoryMinorGridLine) ? resolvedCategoryMinorGridLine : null;
            OfficeColor? valueMinorGridLine = TryGetImageExportMinorGridLineColor(valueAxis, workbookPart, out OfficeColor resolvedValueMinorGridLine) ? resolvedValueMinorGridLine : null;
            double? categoryAxisLineWidth = TryGetImageExportAxisLineWidth(categoryAxis, out double resolvedCategoryAxisLineWidth) ? resolvedCategoryAxisLineWidth : null;
            double? valueAxisLineWidth = TryGetImageExportAxisLineWidth(valueAxis, out double resolvedValueAxisLineWidth) ? resolvedValueAxisLineWidth : null;
            double? categoryGridLineWidth = TryGetImageExportGridLineWidth(categoryAxis, out double resolvedCategoryGridLineWidth) ? resolvedCategoryGridLineWidth : null;
            double? valueGridLineWidth = TryGetImageExportGridLineWidth(valueAxis, out double resolvedValueGridLineWidth) ? resolvedValueGridLineWidth : null;
            double? categoryMinorGridLineWidth = TryGetImageExportMinorGridLineWidth(categoryAxis, out double resolvedCategoryMinorGridLineWidth) ? resolvedCategoryMinorGridLineWidth : null;
            double? valueMinorGridLineWidth = TryGetImageExportMinorGridLineWidth(valueAxis, out double resolvedValueMinorGridLineWidth) ? resolvedValueMinorGridLineWidth : null;
            OfficeStrokeDashStyle? categoryAxisLineDashStyle = TryGetImageExportAxisLineDashStyle(categoryAxis, out OfficeStrokeDashStyle resolvedCategoryAxisLineDashStyle) ? resolvedCategoryAxisLineDashStyle : null;
            OfficeStrokeDashStyle? valueAxisLineDashStyle = TryGetImageExportAxisLineDashStyle(valueAxis, out OfficeStrokeDashStyle resolvedValueAxisLineDashStyle) ? resolvedValueAxisLineDashStyle : null;
            OfficeStrokeDashStyle? categoryGridLineDashStyle = TryGetImageExportGridLineDashStyle(categoryAxis, out OfficeStrokeDashStyle resolvedCategoryGridLineDashStyle) ? resolvedCategoryGridLineDashStyle : null;
            OfficeStrokeDashStyle? valueGridLineDashStyle = TryGetImageExportGridLineDashStyle(valueAxis, out OfficeStrokeDashStyle resolvedValueGridLineDashStyle) ? resolvedValueGridLineDashStyle : null;
            OfficeStrokeDashStyle? categoryMinorGridLineDashStyle = TryGetImageExportMinorGridLineDashStyle(categoryAxis, out OfficeStrokeDashStyle resolvedCategoryMinorGridLineDashStyle) ? resolvedCategoryMinorGridLineDashStyle : null;
            OfficeStrokeDashStyle? valueMinorGridLineDashStyle = TryGetImageExportMinorGridLineDashStyle(valueAxis, out OfficeStrokeDashStyle resolvedValueMinorGridLineDashStyle) ? resolvedValueMinorGridLineDashStyle : null;
            OfficeColor? titleColor = TryGetImageExportTitleColor(chart, workbookPart, out OfficeColor resolvedTitleColor) ? resolvedTitleColor : null;
            string? titleFontFamily = TryGetImageExportTitleFontFamily(chart, workbookPart, out string? resolvedTitleFontFamily) ? resolvedTitleFontFamily : null;
            double? titleFontSize = TryGetImageExportTitleFontSize(chart, workbookPart, out double resolvedTitleFontSize) ? resolvedTitleFontSize : null;
            OfficeFontStyle? titleFontStyle = TryGetImageExportTitleFontStyle(chart, workbookPart, out OfficeFontStyle resolvedTitleFontStyle) ? resolvedTitleFontStyle : null;
            OfficeColor? legendTextColor = TryGetImageExportLegendTextColor(chart, workbookPart, out OfficeColor resolvedLegendTextColor) ? resolvedLegendTextColor : null;
            OfficeColor? dataLabelTextColor = TryGetImageExportDataLabelTextColor(plotArea, workbookPart, out OfficeColor resolvedDataLabelTextColor) ? resolvedDataLabelTextColor : null;
            OfficeColor? dataLabelFill = TryGetImageExportDataLabelFillColor(plotArea, workbookPart, out OfficeColor resolvedDataLabelFill) ? resolvedDataLabelFill : null;
            OfficeColor? dataLabelLine = TryGetImageExportDataLabelLineColor(plotArea, workbookPart, out OfficeColor resolvedDataLabelLine) ? resolvedDataLabelLine : null;
            double? dataLabelLineWidth = TryGetImageExportDataLabelLineWidth(plotArea, out double resolvedDataLabelLineWidth) ? resolvedDataLabelLineWidth : null;
            OfficeStrokeDashStyle? dataLabelLineDashStyle = TryGetImageExportDataLabelLineDashStyle(plotArea, out OfficeStrokeDashStyle resolvedDataLabelLineDashStyle) ? resolvedDataLabelLineDashStyle : null;
            OfficeColor? mutedTextColor = TryGetImageExportAxisTextColor(plotArea, workbookPart, out OfficeColor resolvedMutedTextColor) ? resolvedMutedTextColor : null;
            OfficeColor? axisTitleColor = TryGetImageExportAxisTitleTextColor(plotArea, workbookPart, out OfficeColor resolvedAxisTitleColor) ? resolvedAxisTitleColor : null;
            bool hasNoChartFill = HasNoFill(chartAreaProperties);
            bool hasNoChartLine = HasNoLine(chartAreaProperties);
            bool hasGridLineVisibility = plotArea.Exists && HasImageExportCartesianAxis(plotArea);
            bool showCategoryGridLines = HasImageExportMajorGridlines(categoryAxis);
            bool showValueGridLines = HasImageExportMajorGridlines(valueAxis);
            bool showCategoryMinorGridLines = HasImageExportMinorGridlines(categoryAxis);
            bool showValueMinorGridLines = HasImageExportMinorGridlines(valueAxis);
            OfficeXmlNode radialSeries = ChartType is OfficeChartType.Pie or OfficeChartType.Doughnut
                ? FirstImageExportPieSeries(plotArea)
                : ImageExportPointColoredSeries(plotArea);
            OfficeColor[]? radialPalette = radialSeries.IsNone ? null :
                ImageExportReadRadialPalette(ChartPartRef(), radialSeries,
                    radialPointCount, ImageExportColorScheme(workbookPart));
            if (radialPalette == null && chartFill == null &&
                chartLine == null &&
                plotFill == null &&
                plotLine == null &&
                chartLineWidth == null &&
                plotLineWidth == null &&
                chartLineDashStyle == null &&
                plotLineDashStyle == null &&
                categoryAxisLine == null &&
                valueAxisLine == null &&
                categoryGridLine == null &&
                valueGridLine == null &&
                categoryMinorGridLine == null &&
                valueMinorGridLine == null &&
                categoryAxisLineWidth == null &&
                valueAxisLineWidth == null &&
                categoryGridLineWidth == null &&
                valueGridLineWidth == null &&
                categoryMinorGridLineWidth == null &&
                valueMinorGridLineWidth == null &&
                categoryAxisLineDashStyle == null &&
                valueAxisLineDashStyle == null &&
                categoryGridLineDashStyle == null &&
                valueGridLineDashStyle == null &&
                categoryMinorGridLineDashStyle == null &&
                valueMinorGridLineDashStyle == null &&
                titleColor == null &&
                titleFontFamily == null &&
                titleFontSize == null &&
                titleFontStyle == null &&
                legendTextColor == null &&
                dataLabelTextColor == null &&
                dataLabelFill == null &&
                dataLabelLine == null &&
                dataLabelLineWidth == null &&
                dataLabelLineDashStyle == null &&
                mutedTextColor == null &&
                axisTitleColor == null &&
                !hasNoChartFill &&
                !hasNoChartLine &&
                !hasGridLineVisibility) {
                return null;
            }

            var style = new OfficeChartStyle(palette: radialPalette,
                showBackground: !hasNoChartFill,
                backgroundColor: chartFill,
                borderColor: chartLine,
                legendTextColor: legendTextColor,
                dataLabelTextColor: dataLabelTextColor,
                dataLabelFillColor: dataLabelFill,
                dataLabelBorderColor: dataLabelLine,
                dataLabelBorderWidth: dataLabelLineWidth,
                dataLabelBorderDashStyle: dataLabelLineDashStyle,
                mutedTextColor: mutedTextColor,
                axisTitleColor: axisTitleColor,
                titleColor: titleColor,
                titleFontFamily: titleFontFamily,
                titleFontSize: titleFontSize,
                titleFontStyle: titleFontStyle,
                plotAreaBackgroundColor: plotFill,
                plotAreaBorderColor: plotLine,
                chartBorderWidth: chartLineWidth,
                plotAreaBorderWidth: plotLineWidth,
                chartBorderDashStyle: chartLineDashStyle,
                plotAreaBorderDashStyle: plotLineDashStyle,
                showGridLines: showValueGridLines,
                categoryAxisColor: categoryAxisLine,
                valueAxisColor: valueAxisLine,
                categoryAxisLineWidth: categoryAxisLineWidth,
                valueAxisLineWidth: valueAxisLineWidth,
                categoryAxisLineDashStyle: categoryAxisLineDashStyle,
                valueAxisLineDashStyle: valueAxisLineDashStyle,
                categoryGridLineColor: categoryGridLine,
                valueGridLineColor: valueGridLine,
                categoryGridLineWidth: categoryGridLineWidth,
                valueGridLineWidth: valueGridLineWidth,
                categoryGridLineDashStyle: categoryGridLineDashStyle,
                valueGridLineDashStyle: valueGridLineDashStyle,
                showCategoryGridLines: showCategoryGridLines,
                showValueGridLines: showValueGridLines,
                categoryMinorGridLineColor: categoryMinorGridLine,
                valueMinorGridLineColor: valueMinorGridLine,
                categoryMinorGridLineWidth: categoryMinorGridLineWidth,
                valueMinorGridLineWidth: valueMinorGridLineWidth,
                categoryMinorGridLineDashStyle: categoryMinorGridLineDashStyle,
                valueMinorGridLineDashStyle: valueMinorGridLineDashStyle,
                showCategoryMinorGridLines: showCategoryMinorGridLines,
                showValueMinorGridLines: showValueMinorGridLines,
                showBorder: !hasNoChartLine);

            return style;
        }

        private OfficeChartLayout? CreateImageExportLayout() {
            OfficeXmlNode chart = ChartSpaceNode().Child(C, "chart");
            OfficeXmlNode plotArea = chart.Child(C, "plotArea");
            if (chart.IsNone || plotArea.IsNone) {
                return null;
            }

            OfficeXmlNode[] visibleLabels = plotArea.Descendants(C, "dLbls")
                .Where(labels => IsImageExportSchemaPlaced(labels) && HasAnyVisibleDataLabelPart(labels)).ToArray();
            OfficeXmlNode dataLabels = visibleLabels.FirstOrDefault();
            int[]? radialLabelSeriesIndexes = null;
            OfficeXmlNode radialChart = plotArea.Child(C, "pieChart");
            if (radialChart.IsNone) radialChart = plotArea.Child(C, "doughnutChart");
            OfficeXmlNode chartLevelRadialLabels = radialChart.Child(C, "dLbls");
            if (radialChart.Exists &&
                (chartLevelRadialLabels.IsNone || !HasAnyVisibleDataLabelPart(chartLevelRadialLabels))) {
                OfficeXmlNode[] nativeSeries = radialChart.Children(C, "ser").ToList()
                    .OrderBy(series => series.Child(C, "idx").UInt32Val() ?? uint.MaxValue).ToArray();
                int[] labeled = nativeSeries.Select((series, index) => (series, index))
                    .Where(item => item.series.Child(C, "dLbls") is { Exists: true } labels &&
                        HasAnyVisibleDataLabelPart(labels))
                    .Select(item => item.index).ToArray();
                if (labeled.Length > 0 && labeled.Length < nativeSeries.Length)
                    radialLabelSeriesIndexes = labeled;
            }
            byte[]? firstLabelsXml = radialChart.Exists && visibleLabels.Length > 1 ? dataLabels.OuterXml() : null;
            if (radialChart.Exists && visibleLabels.Skip(1).Any(labels =>
                !labels.OuterXml().AsSpan().SequenceEqual(firstLabelsXml)))
                OfficeIMO.OpenXml.Internal.OfficeOpenXmlChartFidelityScope.Approximate(
                    "Different radial data-label layouts cannot be projected by one shared layout.",
                    "Radial data labels with different layouts per series are drawn with the first labelled series' layout.");
            OfficeXmlNode legend = chart.Child(C, "legend");
            OfficeXmlNode title = chart.Child(C, "title");
            OfficeChartValueAxisLayout? secondaryValueAxis = ReadImageExportSecondaryValueAxis(plotArea);
            OfficeXmlNode categoryAxis = ResolveImageExportCategoryAxis(plotArea);
            OfficeXmlNode valueAxis = ResolveImageExportValueAxis(plotArea);
            string? categoryAxisTitle = GetAxisTitleText(categoryAxis.Child(C, "title"));
            string? valueAxisTitle = GetAxisTitleText(valueAxis.Child(C, "title"));
            string? horizontalAxisNumberFormat = GetImageExportHorizontalAxisNumberFormat(plotArea, categoryAxis, valueAxis);
            string? verticalAxisNumberFormat = GetImageExportVerticalAxisNumberFormat(plotArea, categoryAxis, valueAxis);
            string? categoryAxisNumberFormat = GetImageExportCategoryAxisNumberFormat(categoryAxis);
            (double? Divisor, string? Label) horizontalAxisDisplayUnit = GetImageExportHorizontalAxisDisplayUnit(plotArea, categoryAxis, valueAxis);
            (double? Divisor, string? Label) verticalAxisDisplayUnit = GetImageExportVerticalAxisDisplayUnit(plotArea, categoryAxis, valueAxis);
            (double? Minimum, double? Maximum, double? MajorUnit, double? MinorUnit) horizontalAxisScale = GetImageExportHorizontalAxisScale(plotArea, categoryAxis, valueAxis);
            (double? Minimum, double? Maximum, double? MajorUnit, double? MinorUnit) verticalAxisScale = GetImageExportVerticalAxisScale(plotArea, categoryAxis, valueAxis);
            OfficeChartAxisTickMark horizontalAxisMajorTickMark = GetImageExportHorizontalAxisMajorTickMark(plotArea, categoryAxis, valueAxis);
            OfficeChartAxisTickMark verticalAxisMajorTickMark = GetImageExportVerticalAxisMajorTickMark(plotArea, categoryAxis, valueAxis);
            OfficeChartAxisTickMark horizontalAxisMinorTickMark = GetImageExportHorizontalAxisMinorTickMark(plotArea, categoryAxis, valueAxis);
            OfficeChartAxisTickMark verticalAxisMinorTickMark = GetImageExportVerticalAxisMinorTickMark(plotArea, categoryAxis, valueAxis);
            OfficeChartAxisTickLabelPosition horizontalAxisTickLabelPosition = GetImageExportHorizontalAxisTickLabelPosition(plotArea, categoryAxis, valueAxis);
            OfficeChartAxisTickLabelPosition verticalAxisTickLabelPosition = GetImageExportVerticalAxisTickLabelPosition(plotArea, categoryAxis, valueAxis);
            bool reverseCategoryAxis = GetImageExportReverseCategoryAxis(categoryAxis);
            bool categoryAxisOrientationSpecified = HasImageExportCategoryAxisOrientation(categoryAxis);
            double? legendFontSize = TryGetImageExportLegendFontSize(chart, out double resolvedLegendFontSize) ? resolvedLegendFontSize : null;
            string? legendFontFamily = TryGetImageExportLegendFontFamily(chart, out string? resolvedLegendFontFamily) ? resolvedLegendFontFamily : null;
            OfficeFontStyle? legendFontStyle = TryGetImageExportLegendFontStyle(chart, out OfficeFontStyle resolvedLegendFontStyle) ? resolvedLegendFontStyle : null;
            double? dataLabelFontSize = TryGetImageExportDataLabelFontSize(plotArea, out double resolvedDataLabelFontSize) ? resolvedDataLabelFontSize : null;
            string? dataLabelFontFamily = TryGetImageExportDataLabelFontFamily(plotArea, out string? resolvedDataLabelFontFamily) ? resolvedDataLabelFontFamily : null;
            OfficeFontStyle? dataLabelFontStyle = TryGetImageExportDataLabelFontStyle(plotArea, out OfficeFontStyle resolvedDataLabelFontStyle) ? resolvedDataLabelFontStyle : null;
            double? axisLabelFontSize = TryGetImageExportAxisLabelFontSize(plotArea, out double resolvedAxisLabelFontSize) ? resolvedAxisLabelFontSize : null;
            double? axisTitleFontSize = TryGetImageExportAxisTitleFontSize(plotArea, out double resolvedAxisTitleFontSize) ? resolvedAxisTitleFontSize : null;
            string? axisTextFontFamily = TryGetImageExportAxisTextFontFamily(plotArea, out string? resolvedAxisTextFontFamily) ? resolvedAxisTextFontFamily : null;
            string? axisTitleFontFamily = TryGetImageExportAxisTitleFontFamily(plotArea, out string? resolvedAxisTitleFontFamily) ? resolvedAxisTitleFontFamily : null;
            OfficeFontStyle? axisTextFontStyle = TryGetImageExportAxisTextFontStyle(plotArea, out OfficeFontStyle resolvedAxisTextFontStyle) ? resolvedAxisTextFontStyle : null;
            OfficeFontStyle? axisTitleFontStyle = TryGetImageExportAxisTitleFontStyle(plotArea, out OfficeFontStyle resolvedAxisTitleFontStyle) ? resolvedAxisTitleFontStyle : null;
            bool showCategoryAxis = IsImageExportAxisVisible(categoryAxis);
            bool showValueAxis = IsImageExportAxisVisible(valueAxis);
            bool showCategoryAxisLine = IsImageExportAxisLineVisible(categoryAxis);
            bool showValueAxisLine = IsImageExportAxisLineVisible(valueAxis);
            bool showCategoryAxisLabels = IsImageExportAxisLabelsVisible(categoryAxis);
            bool showValueAxisLabels = IsImageExportAxisLabelsVisible(valueAxis);
            string? scatterStyle = GetImageExportScatterStyle(plotArea);
            bool connectScatterPoints = GetImageExportConnectScatterPoints(scatterStyle);
            bool hasAxisVisibility = !showCategoryAxis || !showValueAxis;
            bool hasAxisLineVisibility = !showCategoryAxisLine || !showValueAxisLine;
            bool hasAxisNumberFormat = horizontalAxisNumberFormat != null || verticalAxisNumberFormat != null || categoryAxisNumberFormat != null;
            bool hasAxisDisplayUnit = horizontalAxisDisplayUnit.Divisor != null || verticalAxisDisplayUnit.Divisor != null;
            bool hasAxisScale = horizontalAxisScale.Minimum != null ||
                horizontalAxisScale.Maximum != null ||
                horizontalAxisScale.MajorUnit != null ||
                horizontalAxisScale.MinorUnit != null ||
                verticalAxisScale.Minimum != null ||
                verticalAxisScale.Maximum != null ||
                verticalAxisScale.MajorUnit != null ||
                verticalAxisScale.MinorUnit != null;
            bool hasAxisMajorTickMark = horizontalAxisMajorTickMark != OfficeChartAxisTickMark.None || verticalAxisMajorTickMark != OfficeChartAxisTickMark.None;
            bool hasAxisMinorTickMark = horizontalAxisMinorTickMark != OfficeChartAxisTickMark.None || verticalAxisMinorTickMark != OfficeChartAxisTickMark.None;
            bool hasAxisLabelVisibility = !showCategoryAxisLabels || !showValueAxisLabels;
            bool hasAxisLabelPosition = horizontalAxisTickLabelPosition != OfficeChartAxisTickLabelPosition.NextTo || verticalAxisTickLabelPosition != OfficeChartAxisTickLabelPosition.NextTo;
            bool hasCategoryAxisOrientation = categoryAxisOrientationSpecified;
            bool fillRadarSeries = GetImageExportFillRadarSeries(plotArea);
            bool hasRadarFillLayout = !fillRadarSeries;
            bool hasScatterStyleLayout = !connectScatterPoints;
            bool hasLegendLayout = legend.IsNone ||
                legend.Child(C, "legendPos").Exists ||
                legend.Child(C, "overlay").Exists;
            bool hasTextFont = legendFontSize != null ||
                legendFontFamily != null ||
                legendFontStyle != null ||
                dataLabelFontSize != null ||
                dataLabelFontFamily != null ||
                dataLabelFontStyle != null ||
                axisLabelFontSize != null ||
                axisTitleFontSize != null ||
                axisTextFontFamily != null ||
                axisTitleFontFamily != null ||
                axisTextFontStyle != null ||
                axisTitleFontStyle != null;
            bool hasLayout =
                secondaryValueAxis != null ||
                dataLabels.Exists ||
                hasLegendLayout ||
                title.Child(C, "overlay").Exists ||
                categoryAxisTitle != null ||
                valueAxisTitle != null ||
                hasAxisVisibility ||
                hasAxisLineVisibility ||
                hasAxisNumberFormat ||
                hasAxisDisplayUnit ||
                hasAxisScale ||
                hasAxisMajorTickMark ||
                hasAxisMinorTickMark ||
                hasAxisLabelVisibility ||
                hasAxisLabelPosition ||
                hasCategoryAxisOrientation ||
                hasRadarFillLayout ||
                hasScatterStyleLayout ||
                hasTextFont;
            if (!hasLayout) {
                // Blank cells and axis crossings still apply to an otherwise default layout.
                OfficeChartLayout semantics = ImageExportWithPlotSemantics(OfficeChartLayout.Default, chart);
                return HasDefaultPlotSemantics(semantics) ? null : semantics;
            }

            bool showLegend = legend.Exists;
            OfficeChartLegendPosition legendPosition = MapLegendPosition(legend.Child(C, "legendPos").Val());
            bool overlayLegend = IsEnabled(legend.Child(C, "overlay"));
            bool showDataLabels = dataLabels.Exists;
            OfficeChartLayout layout = new OfficeChartLayout(
                overlayLegend: overlayLegend,
                showLegend: showLegend,
                legendPosition: legendPosition,
                showDataLabels: showDataLabels,
                showDataLabelValues: IsEnabled(dataLabels.Child(C, "showVal")),
                showDataLabelPercentages: IsEnabled(dataLabels.Child(C, "showPercent")),
                showDataLabelCategoryNames: IsEnabled(dataLabels.Child(C, "showCatName")),
                showDataLabelSeriesNames: IsEnabled(dataLabels.Child(C, "showSerName")),
                dataLabelSeriesIndexes: radialLabelSeriesIndexes,
                dataLabelSeparator: dataLabels.ChildText(C, "separator"),
                legendFontSize: legendFontSize,
                legendFontFamily: legendFontFamily,
                axisLabelFontSize: axisLabelFontSize,
                axisTextFontFamily: axisTextFontFamily,
                dataLabelFontSize: dataLabelFontSize,
                dataLabelFontFamily: dataLabelFontFamily,
                legendFontStyle: legendFontStyle,
                axisTextFontStyle: axisTextFontStyle,
                dataLabelFontStyle: dataLabelFontStyle,
                axisTitleFontSize: axisTitleFontSize,
                axisTitleFontFamily: axisTitleFontFamily,
                axisTitleFontStyle: axisTitleFontStyle,
                dataLabelPosition: MapDataLabelPosition(dataLabels.Child(C, "dLblPos").Val()),
                dataLabelNumberFormat: dataLabels.Child(C, "numFmt").Attribute("formatCode"),
                categoryAxisTitle: categoryAxisTitle,
                valueAxisTitle: valueAxisTitle,
                horizontalAxisNumberFormat: horizontalAxisNumberFormat,
                verticalAxisNumberFormat: verticalAxisNumberFormat,
                categoryAxisNumberFormat: categoryAxisNumberFormat,
                horizontalAxisDisplayUnitDivisor: horizontalAxisDisplayUnit.Divisor,
                horizontalAxisDisplayUnitLabel: horizontalAxisDisplayUnit.Label,
                verticalAxisDisplayUnitDivisor: verticalAxisDisplayUnit.Divisor,
                verticalAxisDisplayUnitLabel: verticalAxisDisplayUnit.Label,
                horizontalAxisMinimum: horizontalAxisScale.Minimum,
                horizontalAxisMaximum: horizontalAxisScale.Maximum,
                horizontalAxisMajorUnit: horizontalAxisScale.MajorUnit,
                horizontalAxisMinorUnit: horizontalAxisScale.MinorUnit,
                verticalAxisMinimum: verticalAxisScale.Minimum,
                verticalAxisMaximum: verticalAxisScale.Maximum,
                verticalAxisMajorUnit: verticalAxisScale.MajorUnit,
                verticalAxisMinorUnit: verticalAxisScale.MinorUnit,
                horizontalAxisMajorTickMark: horizontalAxisMajorTickMark,
                verticalAxisMajorTickMark: verticalAxisMajorTickMark,
                horizontalAxisMinorTickMark: horizontalAxisMinorTickMark,
                verticalAxisMinorTickMark: verticalAxisMinorTickMark,
                horizontalAxisTickLabelPosition: horizontalAxisTickLabelPosition,
                verticalAxisTickLabelPosition: verticalAxisTickLabelPosition,
                reverseCategoryAxis: reverseCategoryAxis,
                categoryAxisOrientationSpecified: categoryAxisOrientationSpecified,
                fillRadarSeries: fillRadarSeries,
                showCategoryAxis: showCategoryAxis,
                showValueAxis: showValueAxis,
                showCategoryAxisLine: showCategoryAxisLine,
                showValueAxisLine: showValueAxisLine,
                showCategoryAxisLabels: showCategoryAxisLabels,
                showValueAxisLabels: showValueAxisLabels,
                connectScatterPoints: connectScatterPoints,
                overlayTitle: IsEnabled(title.Child(C, "overlay")))
                .WithSecondaryValueAxis(secondaryValueAxis);
            layout = ImageExportWithCategoryGeometry(layout.WithDataLabelLeaderLines(IsEnabled(dataLabels.Child(C, "showLeaderLines"))),
                plotArea, ImageExportColorScheme(ExcelPartRef.Workbook(_document)));
            return ImageExportWithPlotSemantics(layout, chart);
        }

        private static bool HasDefaultPlotSemantics(OfficeChartLayout layout) =>
            layout.BlankDisplay == OfficeChartLayout.Default.BlankDisplay &&
            layout.HorizontalAxisCrossingPosition == OfficeChartAxisCrossingPosition.AutoZero && layout.HorizontalAxisCrossesAt == null &&
            layout.VerticalAxisCrossingPosition == OfficeChartAxisCrossingPosition.AutoZero && layout.VerticalAxisCrossesAt == null &&
            layout.HorizontalAxisLogBase == null && layout.VerticalAxisLogBase == null &&
            !layout.ReverseHorizontalValueAxis && !layout.ReverseVerticalValueAxis && !layout.VaryColorsByPoint &&
            layout.PlotAreaLayout == null && layout.TitleLayout == null && layout.LegendLayout == null;

        private OfficeChartValueAxisLayout? ReadImageExportSecondaryValueAxis(OfficeXmlNode plotArea) {
            OfficeXmlNode axis = ImageExportResolveSecondaryAxis(plotArea);
            if (axis.IsNone) return null;
            string? numberFormat = GetImageExportAxisNumberFormat(axis);
            // Qualify the format before the axis layout is built: an axis format the renderer cannot reproduce falls
            // back to General instead of dropping the chart.
            if (numberFormat != null && !IsSimpleSupportedImageExportAxisNumberFormat(numberFormat)) {
                OfficeIMO.OpenXml.Internal.OfficeOpenXmlChartFidelityScope.Record(
                    "The secondary axis number format '" + numberFormat + "' is drawn as General.");
                numberFormat = null;
            }
            OfficeChartValueAxisLayout? native = ImageExportReadSecondaryAxis(
                plotArea, () => numberFormat ?? "General");
            if (native == null) return null;
            var resolved = new OfficeChartValueAxisLayout(native.Minimum, native.Maximum, native.MajorUnit,
                native.MinorUnit, numberFormat, native.MajorTickMark,
                native.MinorTickMark).WithLogarithmicScale(native.LogBase).WithReversedScale(native.IsReversed);
            return native.Title == null ? resolved : resolved.WithTitle(native.Title);
        }

        private static ExcelChartData ApplyImageExportSeriesStyles(ExcelPartRef chartPart, ExcelChartData data, ExcelPartRef workbookPart,
            Func<OfficeXmlNode, IReadOnlyList<double>, IReadOnlyList<double>>? selectVisibleAmounts = null) {
            OfficeXmlNode plotArea = chartPart.Root
                .Child(C, "chart")
                .Child(C, "plotArea");
            if (plotArea.IsNone || data.Series.Count == 0) {
                return data;
            }

            Dictionary<int, ImageExportSeriesStyle> styles = GetImageExportSeriesStyles(plotArea, data, workbookPart, selectVisibleAmounts);
            if (styles.Count == 0) {
                return data;
            }

            var series = new List<ExcelChartSeries>(data.Series.Count);
            bool changed = false;
            for (int i = 0; i < data.Series.Count; i++) {
                ExcelChartSeries current = data.Series[i];
                if (styles.TryGetValue(i, out ImageExportSeriesStyle? style)) {
                    series.Add(current.WithImageExportStyle(style.SeriesColorArgb, style.SeriesLineWidth, style.SeriesLineDashStyle, style.PointColorArgb, style.ShowMarkers, style.ConnectLine, style.MarkerSize.HasValue ? OfficeLength.Points(style.MarkerSize.Value) : (OfficeLength?)null, style.MarkerShape, style.MarkerOutlineColorArgb, style.MarkerOutlineWidth, style.AutomaticColor)
                        .WithSeriesStyle(style.SeriesStyle).WithPointStyles(style.PointStyles).WithPointExplosions(style.PointExplosions)
                        .WithTrendline(style.Trendline).WithErrorBars(style.ErrorBars));
                    changed = true;
                } else {
                    series.Add(current);
                }
            }

            return changed ? new ExcelChartData(data.Categories, series) : data;
        }

        private static Dictionary<int, ImageExportSeriesStyle> GetImageExportSeriesStyles(OfficeXmlNode plotArea, ExcelChartData data, ExcelPartRef workbookPart,
            Func<OfficeXmlNode, IReadOnlyList<double>, IReadOnlyList<double>>? selectVisibleAmounts) {
            var styles = new Dictionary<int, ImageExportSeriesStyle>();
            var automaticColors = new OfficeOpenXmlChartAutomaticColors.SeriesResolutionContext();
            OfficeXmlNode scheme = ImageExportColorScheme(workbookPart);
            int seriesOrder = 0;
            foreach (OfficeXmlNode series in plotArea.Descendants(C, "ser")
                         .Where(IsImageExportSeriesElement)
                         .OrderBy(item => item.Parent.Is(C, "bubbleChart")
                             ? item.Child(C, "order").UInt32Val() ?? uint.MaxValue
                             : item.Child(C, "idx").UInt32Val() ?? uint.MaxValue)) {
                int index = seriesOrder++;
                if (index < 0 || index >= data.Series.Count) {
                    continue;
                }

                ImageExportSeriesStyle style = new ImageExportSeriesStyle();

                OfficeXmlNode properties = series.Child(C, "spPr");
                if (properties.Exists) {
                    style.SeriesStyle = ImageExportReadSeriesStyle(series, scheme);
                    style.SeriesColorArgb = GetImageExportSeriesColor(properties, workbookPart);
                    style.SeriesLineWidth = GetImageExportSeriesLineWidth(properties);
                    style.SeriesLineDashStyle = GetImageExportSeriesLineDashStyle(properties);
                    if (HasNoLine(properties)) {
                        style.ConnectLine = false;
                    }
                }

                style.Trendline = ReadImageExportTrendline(series, workbookPart, data.Series[index].Name);
                style.ErrorBars = ImageExportReadErrorBars(series, scheme, 100000,
                    selectVisibleAmounts);
                int valueCount = data.Series[index].Values.Count;
                OfficeXmlNode marker = ImageExportSeriesChild(series, "marker");
                OfficeXmlNode seriesParent = series.Parent;
                string? scatterStyle = seriesParent.Is(C, "scatterChart") ? seriesParent.Child(C, "scatterStyle").Val() : null;
                style.ShowMarkers = seriesParent.Is(C, "bubbleChart") || GetImageExportShowMarkers(marker, scatterStyle);
                style.MarkerSize = GetImageExportMarkerSize(marker);
                style.MarkerShape = GetImageExportMarkerShape(marker);
                style.MarkerOutlineColorArgb = GetImageExportMarkerOutlineColor(marker, workbookPart);
                style.MarkerOutlineWidth = ReadImageExportMarkerOutlineWidth(marker);
                ApplyImageExportAutomaticAppearance(series, marker, scatterStyle, style,
                    scheme, automaticColors);
                string? markerFill = GetImageExportMarkerFillColor(marker, workbookPart);
                style.PointColorArgb = GetImageExportPointColors(series, valueCount, markerFill, workbookPart);
                IReadOnlyList<OfficeXmlNode> pointOverrides = ImageExportBoundedPoints(series);
                style.PointStyles = ImageExportReadPointStyles(pointOverrides, valueCount, scheme);
                if (IsImageExportPieSeries(series)) {
                    if (!ImageExportTryReadExplosions(series, pointOverrides, valueCount,
                        out int[]? explosions)) {
                        OfficeIMO.OpenXml.Internal.OfficeOpenXmlChartFidelityScope.Approximate(
                            "The native pie or doughnut point explosion cannot be rendered.",
                            "Pie or doughnut slice explosions that cannot be reproduced are drawn without separation.");
                        explosions = null;
                    }
                    style.PointExplosions = explosions;
                }
                if (seriesParent.Is(C, "line3DChart")) {
                    var ribbon = ImageExportReadRibbonPaint(series, scheme, automaticColors);
                    style.SeriesStyle = ribbon.Paint;
                    style.SeriesColorArgb = ribbon.Color?.ToRgbHex();
                    style.AutomaticColor = ribbon.Automatic;
                    style.ShowMarkers = false;
                    style.ConnectLine = ribbon.Paint.ShowOutline != false;
                }

                if (style.HasAny && !styles.ContainsKey(index)) {
                    styles.Add(index, style);
                }
            }

            return styles;
        }

        // Excel colours series without a stored colour from the chart style and theme, and gives
        // line, scatter and marked radar series without a stored symbol its automatic markers.
        private static void ApplyImageExportAutomaticAppearance(OfficeXmlNode series, OfficeXmlNode marker,
            string? scatterStyle, ImageExportSeriesStyle style, OfficeXmlNode scheme,
            OfficeOpenXmlChartAutomaticColors.SeriesResolutionContext automaticColors) {
            if (style.SeriesColorArgb == null && !IsImageExportPieSeries(series) &&
                ImageExportResolveAutomaticSeriesColor(series, scheme, automaticColors) is OfficeColor automatic) {
                style.SeriesColorArgb = automatic.ToRgbHex();
                style.AutomaticColor = true;
            }
            // Office draws automatic markers on line series without a stored marker even when the
            // group's c:marker is 0.
            OfficeXmlNode group = series.Parent;
            bool groupMarkers = group.Is(C, "lineChart") ? true
                : group.Is(C, "scatterChart") ? GetImageExportScatterStyleShowMarkers(scatterStyle)
                : group.Is(C, "radarChart") && group.Child(C, "radarStyle").Val() == "marker";
            string? symbol = marker.Child(C, "symbol").Val();
            if (groupMarkers && (symbol == null || symbol == "auto")) {
                style.ShowMarkers = true;
                style.MarkerShape = OfficeOpenXmlChartAutomaticColors.ResolveMarker(series.Child(C, "idx").UInt32Val() ?? 0U);
                style.MarkerSize ??= OfficeOpenXmlChartAutomaticColors.AutomaticMarkerSize;
            }
        }

        private static bool GetImageExportFillRadarSeries(OfficeXmlNode plotArea) {
            string? radarStyle = plotArea.Child(C, "radarChart").Child(C, "radarStyle").Val();
            return radarStyle == null || radarStyle != "standard";
        }

        private static string? GetImageExportScatterStyle(OfficeXmlNode plotArea) =>
            plotArea.Child(C, "scatterChart").Child(C, "scatterStyle").Val();

        private static bool GetImageExportConnectScatterPoints(string? style) =>
            style == null || style != "marker";

        // The first trendline of a series; the renderer draws one fitted line per series.
        private static OfficeChartTrendline? ReadImageExportTrendline(OfficeXmlNode series, ExcelPartRef workbookPart, string seriesName) {
            OfficeXmlNode trendline = ImageExportSeriesChild(series, "trendline");
            if (trendline.IsNone) return null;
            OfficeChartTrendlineType? type = trendline.Child(C, "trendlineType").Val() switch {
                "exp" => OfficeChartTrendlineType.Exponential,
                "linear" => OfficeChartTrendlineType.Linear,
                "log" => OfficeChartTrendlineType.Logarithmic,
                "movingAvg" => OfficeChartTrendlineType.MovingAverage,
                "poly" => OfficeChartTrendlineType.Polynomial,
                "power" => OfficeChartTrendlineType.Power,
                _ => null
            };
            if (type == null) return null;
            int? order = ReadImageExportByteVal(trendline.Child(C, "order"));
            int? period = trendline.Child(C, "period").UInt32Val() is uint averagePeriod && averagePeriod <= int.MaxValue ? (int)averagePeriod : null;
            if (type == OfficeChartTrendlineType.Polynomial && order is null or < 2 or > 6) return null;
            if (type == OfficeChartTrendlineType.MovingAverage && period is null or < 2) return null;
            // Rendering limits are checked only while the chart is projected for export; reading chart data alone
            // never fails on how a trendline would be drawn.
            bool projecting = OfficeIMO.OpenXml.Internal.OfficeOpenXmlChartFidelityScope.IsProjecting;
            if (period > OfficeIMO.OpenXml.Internal.OfficeOpenXmlChartTrendlines.MaximumPeriod) {
                if (projecting)
                    OfficeIMO.OpenXml.Internal.OfficeOpenXmlChartFidelityScope.Approximate("The native trendline type cannot be projected.",
                        "A trendline the renderer cannot fit is left out.");
                return null;
            }
            OfficeXmlNode properties = trendline.Child(C, "spPr");
            OfficeXmlNode outline = properties.Child(A, "ln");
            OfficeColor? color = TryGetSolidLine(properties, workbookPart, out OfficeColor line) ? line : null;
            double? width = TryGetImageExportLineWidth(properties, out double lineWidth) && lineWidth > 0D ? lineWidth : null;
            double? intercept = trendline.Child(C, "intercept").DoubleVal() is double fixedIntercept && !double.IsNaN(fixedIntercept) && !double.IsInfinity(fixedIntercept)
                ? fixedIntercept : null;
            var result = new OfficeChartTrendline(type.Value, order, period,
                displayEquation: trendline.Child(C, "dispEq").BoolVal() == true,
                displayRSquared: trendline.Child(C, "dispRSqr").BoolVal() == true,
                color: color, width: width.HasValue ? OfficeLength.Points(width.Value) : null, intercept: intercept)
                .WithNativeDetails(trendline.Child(C, "forward").DoubleVal() ?? 0D, trendline.Child(C, "backward").DoubleVal() ?? 0D,
                    ImageExportReadTrendlineLegendName(series, trendline, type.Value, order, period, seriesName));
            // Same test as OfficeOpenXmlChartTrendlines.HidesLine.
            if (outline.Child(A, "noFill").Exists) return result.WithHiddenLine();
            return result.WithDashStyle(ImageExportReadTrendlineDash(outline, projecting));
        }

        private static string? GetImageExportSeriesColor(OfficeXmlNode properties, ExcelPartRef workbookPart) {
            if (TryGetSolidFill(properties, workbookPart, out OfficeColor fill)) {
                return fill.ToRgbHex();
            }

            if (TryGetSolidLine(properties, workbookPart, out OfficeColor line)) {
                return line.ToRgbHex();
            }

            return null;
        }

        private static double? GetImageExportSeriesLineWidth(OfficeXmlNode properties) =>
            TryGetImageExportLineWidth(properties, out double width) ? width : null;

        private static OfficeStrokeDashStyle? GetImageExportSeriesLineDashStyle(OfficeXmlNode properties) =>
            TryGetLineDashStyle(properties, out OfficeStrokeDashStyle dashStyle) ? dashStyle : null;

        private static bool TryGetImageExportAxisLineColor(OfficeXmlNode axis, ExcelPartRef workbookPart, out OfficeColor color) {
            color = default;
            if (axis.IsNone) {
                return false;
            }

            OfficeXmlNode properties = ImageExportAxisShapeProperties(axis);
            return TryGetSolidLine(properties, workbookPart, out color);
        }

        private static bool TryGetImageExportAxisLineWidth(OfficeXmlNode axis, out double width) {
            width = default;
            if (axis.IsNone) {
                return false;
            }

            OfficeXmlNode properties = ImageExportAxisShapeProperties(axis);
            return TryGetImageExportLineWidth(properties, out width);
        }

        private static bool TryGetImageExportAxisLineDashStyle(OfficeXmlNode axis, out OfficeStrokeDashStyle dashStyle) {
            dashStyle = default;
            if (axis.IsNone) {
                return false;
            }

            OfficeXmlNode properties = ImageExportAxisShapeProperties(axis);
            return TryGetLineDashStyle(properties, out dashStyle);
        }

        private static bool TryGetImageExportGridLineColor(OfficeXmlNode axis, ExcelPartRef workbookPart, out OfficeColor color) {
            color = default;
            return TryGetImageExportGridLineColorCore(axis.Child(C, "majorGridlines"), workbookPart, out color);
        }

        private static bool TryGetImageExportMinorGridLineColor(OfficeXmlNode axis, ExcelPartRef workbookPart, out OfficeColor color) {
            color = default;
            return TryGetImageExportGridLineColorCore(axis.Child(C, "minorGridlines"), workbookPart, out color);
        }

        private static bool TryGetImageExportGridLineColorCore(OfficeXmlNode gridlines, ExcelPartRef workbookPart, out OfficeColor color) {
            color = default;
            OfficeXmlNode properties = gridlines.Child(C, "spPr");
            return TryGetSolidLine(properties, workbookPart, out color);
        }

        private static bool TryGetImageExportGridLineWidth(OfficeXmlNode axis, out double width) {
            width = default;
            return TryGetImageExportGridLineWidthCore(axis.Child(C, "majorGridlines"), out width);
        }

        private static bool TryGetImageExportMinorGridLineWidth(OfficeXmlNode axis, out double width) {
            width = default;
            return TryGetImageExportGridLineWidthCore(axis.Child(C, "minorGridlines"), out width);
        }

        private static bool TryGetImageExportGridLineWidthCore(OfficeXmlNode gridlines, out double width) {
            width = default;
            OfficeXmlNode properties = gridlines.Child(C, "spPr");
            return TryGetImageExportLineWidth(properties, out width);
        }

        private static bool TryGetImageExportGridLineDashStyle(OfficeXmlNode axis, out OfficeStrokeDashStyle dashStyle) {
            dashStyle = default;
            return TryGetImageExportGridLineDashStyleCore(axis.Child(C, "majorGridlines"), out dashStyle);
        }

        private static bool TryGetImageExportMinorGridLineDashStyle(OfficeXmlNode axis, out OfficeStrokeDashStyle dashStyle) {
            dashStyle = default;
            return TryGetImageExportGridLineDashStyleCore(axis.Child(C, "minorGridlines"), out dashStyle);
        }

        private static bool TryGetImageExportGridLineDashStyleCore(OfficeXmlNode gridlines, out OfficeStrokeDashStyle dashStyle) {
            dashStyle = default;
            OfficeXmlNode properties = gridlines.Child(C, "spPr");
            return TryGetLineDashStyle(properties, out dashStyle);
        }

        private static bool HasImageExportCartesianAxis(OfficeXmlNode plotArea) =>
            GetImageExportAxes(plotArea).Count > 0;

        private static bool HasImageExportMajorGridlines(OfficeXmlNode axis) =>
            axis.Child(C, "majorGridlines").Exists;

        private static bool HasImageExportMinorGridlines(OfficeXmlNode axis) =>
            axis.Child(C, "minorGridlines").Exists;

        // The category, value and date axes, in that order.
        private static List<OfficeXmlNode> GetImageExportAxes(OfficeXmlNode plotArea) {
            var axes = new List<OfficeXmlNode>();
            foreach (OfficeXmlNode axis in plotArea.Children(C, "catAx")) axes.Add(axis);
            foreach (OfficeXmlNode axis in plotArea.Children(C, "valAx")) axes.Add(axis);
            foreach (OfficeXmlNode axis in plotArea.Children(C, "dateAx")) axes.Add(axis);
            return axes;
        }

        private static OfficeXmlNode ResolveImageExportCategoryAxis(OfficeXmlNode plotArea) {
            OfficeXmlNode axis = ResolveCategoryAxis(plotArea, OfficeChartAxisGroup.Primary);
            return axis.Exists ? axis : ResolveScatterXAxis(plotArea);
        }

        private static OfficeXmlNode ResolveImageExportValueAxis(OfficeXmlNode plotArea) {
            OfficeXmlNode axis = ResolveValueAxis(plotArea, OfficeChartAxisGroup.Primary);
            return axis.Exists ? axis : ResolveScatterYAxis(plotArea);
        }

        private static bool IsImageExportAxisVisible(OfficeXmlNode axis) =>
            axis.IsNone || !IsEnabled(axis.Child(C, "delete"));

        private static bool IsImageExportAxisLineVisible(OfficeXmlNode axis) =>
            axis.IsNone || (IsImageExportAxisVisible(axis) && !HasNoLine(ImageExportAxisShapeProperties(axis)));

        private static bool IsImageExportAxisLabelsVisible(OfficeXmlNode axis) =>
            axis.IsNone || (IsImageExportAxisVisible(axis) && axis.Child(C, "tickLblPos").Val() != "none");

        private static IReadOnlyList<string?>? GetImageExportPointColors(OfficeXmlNode series, int valueCount, string? markerFill, ExcelPartRef workbookPart) {
            if (valueCount <= 0 && string.IsNullOrWhiteSpace(markerFill)) {
                return null;
            }

            var colors = new string?[valueCount];
            bool any = false;
            if (!string.IsNullOrWhiteSpace(markerFill)) {
                for (int i = 0; i < colors.Length; i++) {
                    colors[i] = markerFill;
                }

                any = colors.Length > 0;
            }

            foreach (OfficeXmlNode point in ImageExportBoundedPoints(series)) {
                uint? rawIndex = point.Child(C, "idx").UInt32Val();
                if (rawIndex == null || rawIndex.Value > int.MaxValue) {
                    continue;
                }

                int index = (int)rawIndex.Value;
                if (index >= colors.Length) {
                    continue;
                }

                OfficeXmlNode properties = point.Child(C, "spPr");
                if (properties.IsNone || !TryGetSolidFill(properties, workbookPart, out OfficeColor color)) {
                    continue;
                }

                colors[index] = color.ToRgbHex();
                any = true;
            }

            return any ? colors : null;
        }

        private static bool GetImageExportShowMarkers(OfficeXmlNode marker, string? scatterStyle) {
            if (marker.IsNone) {
                return scatterStyle != null && GetImageExportScatterStyleShowMarkers(scatterStyle);
            }

            string? symbol = marker.Child(C, "symbol").Val();
            return symbol == null || symbol != "none";
        }

        private static bool GetImageExportScatterStyleShowMarkers(string? style) {
            if (style == null) {
                return true;
            }

            return style == "marker" ||
                style == "lineMarker" ||
                style == "smoothMarker";
        }

        private static int? GetImageExportMarkerSize(OfficeXmlNode marker) {
            int? size = ReadImageExportByteVal(marker.Child(C, "size"));
            return size.HasValue && size.Value > 0 ? size.Value : null;
        }

        private static OfficeChartMarkerShape? GetImageExportMarkerShape(OfficeXmlNode marker) {
            if (marker.IsNone) {
                return null;
            }

            switch (marker.Child(C, "symbol").Val()) {
                case null:
                case "auto":
                case "circle":
                    return OfficeChartMarkerShape.Circle;
                case "square":
                    return OfficeChartMarkerShape.Square;
                case "diamond":
                    return OfficeChartMarkerShape.Diamond;
                case "triangle":
                    return OfficeChartMarkerShape.Triangle;
                case "dash":
                    return OfficeChartMarkerShape.Dash;
                case "dot":
                    return OfficeChartMarkerShape.Dot;
                case "plus":
                    return OfficeChartMarkerShape.Plus;
                case "x":
                    return OfficeChartMarkerShape.X;
                case "star":
                    return OfficeChartMarkerShape.Star;
                default:
                    return null;
            }
        }

        private static string? GetImageExportMarkerFillColor(OfficeXmlNode marker, ExcelPartRef workbookPart) {
            OfficeXmlNode properties = marker.Child(C, "spPr");
            return properties.Exists && TryGetSolidFill(properties, workbookPart, out OfficeColor color)
                ? color.A == 255 ? color.ToRgbHex() : color.ToArgbHex()
                : null;
        }

        private static string? GetImageExportMarkerOutlineColor(OfficeXmlNode marker, ExcelPartRef workbookPart) {
            OfficeXmlNode properties = marker.Child(C, "spPr");
            return properties.Exists && TryGetSolidLine(properties, workbookPart, out OfficeColor color)
                ? color.A == 255 ? color.ToRgbHex() : color.ToArgbHex()
                : null;
        }

        private static double? ReadImageExportMarkerOutlineWidth(OfficeXmlNode marker) {
            OfficeXmlNode outline = marker
                .Child(C, "spPr")
                .Child(A, "ln");
            long? emu = outline.Int32Attribute("w");
            return emu.HasValue && emu.Value > 0
                ? ExcelImageExportLimits.ClampStrokeWidth(emu.Value / 12700D)
                : null;
        }

        private IReadOnlyList<OfficeImageExportDiagnostic> CreateImageExportDiagnostics() {
            OfficeXmlNode chartSpace = ChartSpaceNode();
            if (chartSpace.IsNone) {
                return Array.Empty<OfficeImageExportDiagnostic>();
            }

            var diagnostics = new List<OfficeImageExportDiagnostic>();
            string source = GetImageExportSource();
            ExcelPartRef workbookPart = ExcelPartRef.Workbook(_document);
            // Every trendline type, its forecasts and its legend entry are drawn; only the first trendline
            // of a series is, so further trendlines still warn.
            if (chartSpace.Descendants(C, "trendline").Any(trendline => IsImageExportSchemaPlaced(trendline) &&
                trendline.Parent.Children(C, "trendline").Count > 1)) {
                diagnostics.Add(ExcelImageExportDiagnosticClassifier.Create(
                    OfficeImageExportDiagnosticSeverity.Warning,
                    ExcelImageExportDiagnosticCodes.ChartTrendlineUnsupported,
                    "Worksheet chart trendlines after the first on a series are not rendered by the shared image renderer.",
                    source));
            }

            if (chartSpace.Descendants(C, "dLbl").Any(IsImageExportSchemaPlaced)) {
                diagnostics.Add(ExcelImageExportDiagnosticClassifier.Create(
                    OfficeImageExportDiagnosticSeverity.Warning,
                    ExcelImageExportDiagnosticCodes.ChartDataLabelPointOverridesApproximated,
                    "Worksheet chart point-level data-label overrides are approximated as chart-level data labels in image export.",
                    source));
            }

            // Outside-end and best-fit pie labels draw their leader lines in the default line style.
            if (chartSpace.Descendants(C, "leaderLines").Any(lines => IsImageExportSchemaPlaced(lines) && HasUnsupportedImageExportLeaderLines(lines))) {
                diagnostics.Add(ExcelImageExportDiagnosticClassifier.Create(
                    OfficeImageExportDiagnosticSeverity.Warning,
                    ExcelImageExportDiagnosticCodes.ChartDataLabelLeaderLinesUnsupported,
                    "Worksheet chart data-label leader lines are not rendered by the shared image renderer yet.",
                    source));
            }

            if (HasUnsupportedImageExportGridlineStyle(chartSpace, workbookPart)) {
                diagnostics.Add(ExcelImageExportDiagnosticClassifier.Create(
                    OfficeImageExportDiagnosticSeverity.Warning,
                    ExcelImageExportDiagnosticCodes.ChartGridlineStyleApproximation,
                    "Worksheet chart gridline styling is only partially represented by the shared image renderer.",
                    source));
            }

            if (HasUnsupportedImageExportAxisStyle(chartSpace, workbookPart)) {
                diagnostics.Add(ExcelImageExportDiagnosticClassifier.Create(
                    OfficeImageExportDiagnosticSeverity.Warning,
                    ExcelImageExportDiagnosticCodes.ChartAxisStyleApproximation,
                    "Worksheet chart axis line styling is only partially represented by the shared image renderer.",
                    source));
            }

            if (HasUnsupportedImageExportAxisTickLabelPosition(chartSpace)) {
                diagnostics.Add(ExcelImageExportDiagnosticClassifier.Create(
                    OfficeImageExportDiagnosticSeverity.Warning,
                    ExcelImageExportDiagnosticCodes.ChartAxisTickLabelPositionApproximation,
                    "Worksheet chart axis tick-label positions outside next-to, high, low, or none are approximated by the shared image renderer.",
                    source));
            }

            if (HasApproximatedImageExportAxisMinorTickMarks(chartSpace)) {
                diagnostics.Add(ExcelImageExportDiagnosticClassifier.Create(
                    OfficeImageExportDiagnosticSeverity.Warning,
                    ExcelImageExportDiagnosticCodes.ChartAxisMinorTickMarkPlacementApproximation,
                    "Worksheet chart minor axis tick mark placement is approximated by the shared image renderer.",
                    source));
            }

            if (HasUnsupportedImageExportAxisCrossing(chartSpace)) {
                diagnostics.Add(ExcelImageExportDiagnosticClassifier.Create(
                    OfficeImageExportDiagnosticSeverity.Warning,
                    ExcelImageExportDiagnosticCodes.ChartAxisCrossingApproximation,
                    "Worksheet chart custom axis crossing is only partially positioned by the shared image renderer.",
                    source));
            }

            if (HasUnsupportedImageExportAxisScale(chartSpace)) {
                diagnostics.Add(ExcelImageExportDiagnosticClassifier.Create(
                    OfficeImageExportDiagnosticSeverity.Warning,
                    ExcelImageExportDiagnosticCodes.ChartAxisScaleApproximation,
                    "Worksheet chart custom axis scale, units, or reverse-order settings are not applied by the shared image renderer yet.",
                    source));
            }

            if (HasUnsupportedImageExportAxisNumberFormat(chartSpace)) {
                diagnostics.Add(ExcelImageExportDiagnosticClassifier.Create(
                    OfficeImageExportDiagnosticSeverity.Warning,
                    ExcelImageExportDiagnosticCodes.ChartAxisNumberFormatApproximation,
                    "Worksheet chart axis number formatting is only partially represented by the shared image renderer.",
                    source));
            }

            if (HasUnsupportedImageExportCategoryAxisNumberFormat(chartSpace)) {
                diagnostics.Add(ExcelImageExportDiagnosticClassifier.Create(
                    OfficeImageExportDiagnosticSeverity.Warning,
                    ExcelImageExportDiagnosticCodes.ChartCategoryAxisNumberFormatUnsupported,
                    "Worksheet chart category or date axis number formats are not rendered by the shared image renderer yet.",
                    source));
            }

            if (HasUnsupportedImageExportTextStyle(chartSpace, workbookPart)) {
                diagnostics.Add(ExcelImageExportDiagnosticClassifier.Create(
                    OfficeImageExportDiagnosticSeverity.Warning,
                    ExcelImageExportDiagnosticCodes.ChartTextStyleApproximation,
                    "Worksheet chart text styling is only partially represented by the shared image renderer.",
                    source));
            }

            if (HasUnsupportedImageExportChartAreaStyle(chartSpace, workbookPart)) {
                diagnostics.Add(ExcelImageExportDiagnosticClassifier.Create(
                    OfficeImageExportDiagnosticSeverity.Warning,
                    ExcelImageExportDiagnosticCodes.ChartAreaStyleApproximation,
                    "Worksheet chart or plot area styling is only partially represented by the shared image renderer.",
                    source));
            }

            if (chartSpace.Descendants(C, "spPr").Any(properties => IsImageExportSchemaPlaced(properties) &&
                IsUnsupportedImageExportSeriesOrMarkerShapeProperties(properties, workbookPart))) {
                diagnostics.Add(ExcelImageExportDiagnosticClassifier.Create(
                    OfficeImageExportDiagnosticSeverity.Warning,
                    ExcelImageExportDiagnosticCodes.ChartSeriesStyleApproximation,
                    "Worksheet chart series, marker, or data-label shape styling is only partially represented by the shared image renderer.",
                    source));
            }

            return diagnostics;
        }

        private string GetImageExportSource() =>
            _sheetName + "!" + (string.IsNullOrWhiteSpace(Name) ? ChartType.ToString() : Name);

        // Pie and doughnut labels at the outside end or best fit draw their leader lines in the default line style.
        private static bool HasUnsupportedImageExportLeaderLines(OfficeXmlNode lines) {
            if (lines.FirstChild.Exists || ImageExportHasAttributes(lines)) return true;
            if (lines.Ancestor(C, "pieChart").IsNone && lines.Ancestor(C, "doughnutChart").IsNone) return true;
            OfficeXmlNode labels = lines.Parent;
            if (!labels.Is(C, "dLbls")) return true;
            string? position = labels.Child(C, "dLblPos").Val();
            if (position != null && position != "outEnd" && position != "bestFit") return true;
            OfficeXmlNode showLeaderLines = labels.Child(C, "showLeaderLines");
            return showLeaderLines.IsNone || showLeaderLines.BoolVal() == false;
        }

        private static bool HasAnyVisibleDataLabelPart(OfficeXmlNode labels) =>
            IsEnabled(labels.Child(C, "showVal")) ||
            IsEnabled(labels.Child(C, "showPercent")) ||
            IsEnabled(labels.Child(C, "showCatName")) ||
            IsEnabled(labels.Child(C, "showSerName"));

        private static bool IsEnabled(OfficeXmlNode value) =>
            value.Exists && (value.BoolVal() ?? true);

        private static bool TryGetSolidFill(OfficeXmlNode properties, out OfficeColor color) =>
            TryGetSolidFill(properties, ExcelPartRef.None, out color);

        private static bool TryGetSolidFill(OfficeXmlNode properties, ExcelPartRef workbookPart, out OfficeColor color) {
            string? value = ResolveImageExportSolidFill(properties.Child(A, "solidFill"), workbookPart);
            return TryParseResolvedDrawingColor(value, out color);
        }

        private static bool TryGetSolidLine(OfficeXmlNode properties, out OfficeColor color) =>
            TryGetSolidLine(properties, ExcelPartRef.None, out color);

        private static bool TryGetSolidLine(OfficeXmlNode properties, ExcelPartRef workbookPart, out OfficeColor color) {
            string? value = ResolveImageExportSolidFill(
                properties
                    .Child(A, "ln")
                    .Child(A, "solidFill"),
                workbookPart);
            return TryParseResolvedDrawingColor(value, out color);
        }

        private static bool TryParseResolvedDrawingColor(string? argb, out OfficeColor color) {
            color = default;
            if (string.IsNullOrWhiteSpace(argb)) {
                return false;
            }

            string value = argb!.Trim().TrimStart('#');
            if (value.Length == 8) {
                value = value.Substring(2, 6) + value.Substring(0, 2);
            }

            return OfficeColor.TryParseHex(value, out color);
        }

        private static bool HasNoFill(OfficeXmlNode properties) =>
            properties.Child(A, "noFill").Exists;

        private static bool HasNoLine(OfficeXmlNode properties) =>
            properties.Child(A, "ln").Child(A, "noFill").Exists;

        private static bool TryGetImageExportLineWidth(OfficeXmlNode properties, out double width) {
            width = default;
            long? emu = properties
                .Child(A, "ln")
                .Int32Attribute("w");
            if (!emu.HasValue || emu.Value <= 0) {
                return false;
            }

            width = ExcelImageExportLimits.ClampStrokeWidth(emu.Value / 12700D);
            return true;
        }

        private static bool TryGetLineDashStyle(OfficeXmlNode properties, out OfficeStrokeDashStyle dashStyle) {
            dashStyle = default;
            string? value = properties
                .Child(A, "ln")
                .Child(A, "prstDash")
                .Val();
            if (value == null || value == "solid") {
                return false;
            }

            return TryMapPresetLineDash(value, out dashStyle);
        }

        private static bool TryMapPresetLineDash(string value, out OfficeStrokeDashStyle dashStyle) {
            return OfficeStrokeDashStyleMapper.TryMapOfficePresetDash(GetPresetLineDashToken(value), out dashStyle);
        }

        // The schema tokens the shared mapper knows; solid and anything else map to none.
        private static string? GetPresetLineDashToken(string value) => value switch {
            "dash" or "lgDash" or "sysDash" or "dot" or "sysDot" or "dashDot" or "lgDashDot" or "sysDashDot" or
                "lgDashDotDot" or "sysDashDotDot" => value,
            _ => null
        };

        // A dash value the preset dash mapper cannot draw; solid and a missing value are drawable.
        private static bool IsUndrawableImageExportDash(OfficeXmlNode dash) {
            string? value = dash.Val();
            return value != null && value != "solid" && !TryMapPresetLineDash(value, out _);
        }

        private static bool HasUnsupportedImageExportGridlineStyle(OfficeXmlNode chartSpace, ExcelPartRef workbookPart) {
            foreach (OfficeXmlNode gridlines in chartSpace.Descendants(C, "majorGridlines")) {
                OfficeXmlNode properties = gridlines.Child(C, "spPr");
                if (properties.Exists && !IsSimpleSupportedGridlineShapeProperties(properties, workbookPart)) {
                    return true;
                }
            }

            foreach (OfficeXmlNode gridlines in chartSpace.Descendants(C, "minorGridlines")) {
                OfficeXmlNode properties = gridlines.Child(C, "spPr");
                if (properties.Exists && !IsSimpleSupportedGridlineShapeProperties(properties, workbookPart)) {
                    return true;
                }
            }

            return false;
        }

        private static bool IsSimpleSupportedGridlineShapeProperties(OfficeXmlNode properties, ExcelPartRef workbookPart) {
            if (properties.FirstChild.IsNone) {
                return true;
            }

            foreach (OfficeXmlNode child in properties.Children()) {
                if (child.Is(A, "ln")) {
                    foreach (OfficeXmlNode outlineChild in child.Children()) {
                        if (outlineChild.Is(A, "solidFill")) {
                            if (!TryGetSolidLine(properties, workbookPart, out _)) {
                                return false;
                            }

                            continue;
                        }

                        if (outlineChild.Is(A, "prstDash")) {
                            if (IsUndrawableImageExportDash(outlineChild)) {
                                return false;
                            }

                            continue;
                        }

                        return false;
                    }

                    continue;
                }

                return false;
            }

            return true;
        }

        private static bool HasUnsupportedImageExportAxisStyle(OfficeXmlNode chartSpace, ExcelPartRef workbookPart) {
            foreach (OfficeXmlNode plotArea in chartSpace.Descendants(C, "plotArea")) {
                foreach (OfficeXmlNode axis in GetImageExportAxes(plotArea)) {
                    OfficeXmlNode properties = ImageExportAxisShapeProperties(axis);
                    if (properties.Exists && !IsSimpleSupportedAxisShapeProperties(properties, workbookPart)) {
                        return true;
                    }
                }
            }

            return false;
        }

        private static bool IsSimpleSupportedAxisShapeProperties(OfficeXmlNode properties, ExcelPartRef workbookPart) {
            if (properties.FirstChild.IsNone) {
                return true;
            }

            foreach (OfficeXmlNode child in properties.Children()) {
                if (child.Is(A, "ln")) {
                    foreach (OfficeXmlNode outlineChild in child.Children()) {
                        if (outlineChild.Is(A, "solidFill")) {
                            if (!TryGetSolidLine(properties, workbookPart, out _)) {
                                return false;
                            }

                            continue;
                        }

                        if (outlineChild.Is(A, "noFill")) {
                            continue;
                        }

                        if (outlineChild.Is(A, "prstDash")) {
                            if (IsUndrawableImageExportDash(outlineChild)) {
                                return false;
                            }

                            continue;
                        }

                        return false;
                    }

                    continue;
                }

                return false;
            }

            return true;
        }

        private static bool HasUnsupportedImageExportAxisTickLabelPosition(OfficeXmlNode chartSpace) {
            foreach (OfficeXmlNode plotArea in chartSpace.Descendants(C, "plotArea")) {
                foreach (OfficeXmlNode axis in GetImageExportAxes(plotArea)) {
                    string? position = axis.Child(C, "tickLblPos").Val();
                    if (position != null &&
                        position != "nextTo" &&
                        position != "none" &&
                        position != "high" &&
                        position != "low") {
                        return true;
                    }
                }
            }

            return false;
        }

        private static bool HasApproximatedImageExportAxisMinorTickMarks(OfficeXmlNode chartSpace) {
            foreach (OfficeXmlNode plotArea in chartSpace.Descendants(C, "plotArea")) {
                foreach (OfficeXmlNode axis in GetImageExportAxes(plotArea)) {
                    if (HasApproximatedImageExportAxisMinorTickMark(axis)) {
                        return true;
                    }
                }
            }

            return false;
        }

        private static bool HasApproximatedImageExportAxisMinorTickMark(OfficeXmlNode axis) {
            string? tickMark = axis.Child(C, "minorTickMark").Val();
            if (tickMark == null || tickMark == "none") {
                return false;
            }

            if (!axis.Is(C, "valAx")) {
                return true;
            }

            double? minorUnit = axis.Child(C, "minorUnit").DoubleVal();
            return !minorUnit.HasValue ||
                   double.IsNaN(minorUnit.Value) ||
                   double.IsInfinity(minorUnit.Value) ||
                   minorUnit.Value <= 0D;
        }

        private static bool HasUnsupportedImageExportAxisCrossing(OfficeXmlNode chartSpace) {
            foreach (OfficeXmlNode plotArea in chartSpace.Descendants(C, "plotArea")) {
                foreach (OfficeXmlNode axis in GetImageExportAxes(plotArea)) {
                    if (axis.Child(C, "crossesAt").Exists) {
                        return true;
                    }

                    string? crosses = axis.Child(C, "crosses").Val();
                    if (crosses != null &&
                        crosses != "autoZero" &&
                        !IsSupportedImageExportAxisCrossing(plotArea, axis, crosses)) {
                        return true;
                    }
                }
            }

            return false;
        }

        private static bool IsSupportedImageExportAxisCrossing(OfficeXmlNode plotArea, OfficeXmlNode axis, string crosses) =>
            crosses == "max" &&
            (axis.Is(C, "valAx") || axis.Is(C, "catAx") || axis.Is(C, "dateAx")) &&
            !HasHorizontalBarChart(plotArea);

        private static bool HasUnsupportedImageExportAxisScale(OfficeXmlNode chartSpace) {
            foreach (OfficeXmlNode plotArea in chartSpace.Descendants(C, "plotArea")) {
                foreach (OfficeXmlNode axis in GetImageExportAxes(plotArea)) {
                    bool valueAxis = axis.Is(C, "valAx");
                    OfficeXmlNode scaling = axis.Child(C, "scaling");
                    if (scaling.Exists && HasUnsupportedImageExportAxisScaling(plotArea, axis, scaling)) {
                        return true;
                    }

                    if (axis.Child(C, "majorUnit").Exists && !valueAxis) {
                        return true;
                    }

                    OfficeXmlNode minorUnit = axis.Child(C, "minorUnit");
                    double? minorUnitValue = minorUnit.DoubleVal();
                    if (minorUnit.Exists &&
                        (!valueAxis ||
                         !minorUnitValue.HasValue ||
                         double.IsNaN(minorUnitValue.Value) ||
                         double.IsInfinity(minorUnitValue.Value) ||
                         minorUnitValue.Value <= 0D)) {
                        return true;
                    }

                    string? crossBetween = axis.Child(C, "crossBetween").Val();
                    if (crossBetween != null && crossBetween != "between") {
                        return true;
                    }
                }
            }

            return false;
        }

        private static bool HasUnsupportedImageExportAxisScaling(OfficeXmlNode plotArea, OfficeXmlNode axis, OfficeXmlNode scaling) {
            // Reversed value and category axes and logarithmic value axes are drawn by the shared renderer.
            bool valueAxis = axis.Is(C, "valAx");
            if (scaling.Child(C, "logBase").Exists && !valueAxis) {
                return true;
            }

            OfficeXmlNode minimum = scaling.Child(C, "min");
            OfficeXmlNode maximum = scaling.Child(C, "max");
            if ((minimum.Exists || maximum.Exists) && !valueAxis) {
                return true;
            }

            double? min = minimum.DoubleVal();
            double? max = maximum.DoubleVal();
            if (min.HasValue && (double.IsNaN(min.Value) || double.IsInfinity(min.Value))) {
                return true;
            }

            if (max.HasValue && (double.IsNaN(max.Value) || double.IsInfinity(max.Value))) {
                return true;
            }

            return min.HasValue && max.HasValue && max.Value <= min.Value;
        }


        private static bool HasUnsupportedImageExportChartAreaStyle(OfficeXmlNode chartSpace, ExcelPartRef workbookPart) {
            OfficeXmlNode chartAreaProperties = chartSpace.Child(C, "spPr");
            if (chartAreaProperties.Exists && !IsSimpleSupportedChartAreaShapeProperties(chartAreaProperties, workbookPart)) {
                return true;
            }

            foreach (OfficeXmlNode plotArea in chartSpace.Descendants(C, "plotArea")) {
                OfficeXmlNode properties = plotArea.Child(C, "spPr");
                if (properties.Exists && !IsSimpleSupportedChartAreaShapeProperties(properties, workbookPart)) {
                    return true;
                }
            }

            return false;
        }

        private static bool IsSimpleSupportedChartAreaShapeProperties(OfficeXmlNode properties, ExcelPartRef workbookPart) {
            if (properties.FirstChild.IsNone) {
                return true;
            }

            foreach (OfficeXmlNode child in properties.Children()) {
                if (child.Is(A, "solidFill")) {
                    if (!TryGetSolidFill(properties, workbookPart, out _)) {
                        return false;
                    }

                    continue;
                }

                if (child.Is(A, "noFill")) {
                    continue;
                }

                if (child.Is(A, "ln")) {
                    if (child.Attribute("w") != null && !TryGetImageExportLineWidth(properties, out _)) {
                        return false;
                    }

                    foreach (OfficeXmlNode outlineChild in child.Children()) {
                        if (outlineChild.Is(A, "solidFill")) {
                            if (!TryGetSolidLine(properties, workbookPart, out _)) {
                                return false;
                            }

                            continue;
                        }

                        if (outlineChild.Is(A, "noFill")) {
                            continue;
                        }

                        if (outlineChild.Is(A, "prstDash")) {
                            if (IsUndrawableImageExportDash(outlineChild)) {
                                return false;
                            }

                            continue;
                        }

                        return false;
                    }

                    continue;
                }

                return false;
            }

            return true;
        }

        // The c:spPr of a series, data point, marker, data label or trendline.
        private static bool IsSeriesOrMarkerShapeProperties(OfficeXmlNode properties) {
            OfficeXmlNode parent = properties.Parent;
            return IsImageExportSeriesElement(parent)
                || parent.Is(C, "dPt")
                || parent.Is(C, "marker")
                || parent.Is(C, "dLbls")
                || parent.Is(C, "dLbl")
                || parent.Is(C, "trendline");
        }

        private static bool IsUnsupportedImageExportSeriesOrMarkerShapeProperties(OfficeXmlNode properties, ExcelPartRef workbookPart) {
            if (!IsSeriesOrMarkerShapeProperties(properties)) {
                return false;
            }

            OfficeXmlNode parent = properties.Parent;
            if (parent.Is(C, "marker")) {
                return !IsSimpleSupportedMarker(parent, workbookPart);
            }

            if (parent.Is(C, "dPt")) {
                return !IsSimpleSupportedDataPoint(parent, workbookPart);
            }

            if (parent.Is(C, "dLbls")) {
                return !IsSimpleSupportedSeriesShapeProperties(properties, workbookPart);
            }

            if (parent.Is(C, "dLbl") || parent.Is(C, "trendline")) {
                return true;
            }

            return !IsSimpleSupportedSeriesShapeProperties(properties, workbookPart);
        }

        private static bool IsSimpleSupportedSeriesShapeProperties(OfficeXmlNode properties, ExcelPartRef workbookPart) {
            if (properties.FirstChild.IsNone) {
                return true;
            }

            foreach (OfficeXmlNode child in properties.Children()) {
                if (child.Is(A, "solidFill")) {
                    if (!TryGetSolidFill(properties, workbookPart, out _)) {
                        return false;
                    }

                    continue;
                }

                if (child.Is(A, "ln")) {
                    if (child.Attribute("w") != null && !TryGetImageExportLineWidth(properties, out _)) {
                        return false;
                    }

                    foreach (OfficeXmlNode outlineChild in child.Children()) {
                        if (outlineChild.Is(A, "solidFill")) {
                            if (!TryGetSolidLine(properties, workbookPart, out _)) {
                                return false;
                            }

                            continue;
                        }

                        if (outlineChild.Is(A, "prstDash")) {
                            if (IsUndrawableImageExportDash(outlineChild)) {
                                return false;
                            }

                            continue;
                        }

                        return false;
                    }

                    continue;
                }

                return false;
            }

            return true;
        }

        private static bool IsSimpleSupportedMarker(OfficeXmlNode marker, ExcelPartRef workbookPart) {
            string? symbol = marker.Child(C, "symbol").Val();
            if (symbol != null &&
                symbol != "circle" &&
                symbol != "auto" &&
                symbol != "square" &&
                symbol != "diamond" &&
                symbol != "triangle" &&
                symbol != "dash" &&
                symbol != "dot" &&
                symbol != "plus" &&
                symbol != "x" &&
                symbol != "star" &&
                symbol != "none") {
                return false;
            }

            OfficeXmlNode properties = marker.Child(C, "spPr");
            if (properties.IsNone) return true;
            return marker.Parent.Is(C, "dPt")
                ? ImageExportIsSupportedPointStyle(properties, ImageExportColorScheme(workbookPart))
                : IsSimpleSupportedMarkerShapeProperties(properties, workbookPart);
        }

        private static bool IsSimpleSupportedMarkerShapeProperties(OfficeXmlNode properties, ExcelPartRef workbookPart) {
            if (properties.FirstChild.IsNone) {
                return true;
            }

            foreach (OfficeXmlNode child in properties.Children()) {
                if (child.Is(A, "solidFill")) {
                    if (!TryGetSolidFill(properties, workbookPart, out _)) {
                        return false;
                    }

                    continue;
                }

                if (child.Is(A, "ln")) {
                    foreach (OfficeXmlNode outlineChild in child.Children()) {
                        if (outlineChild.Is(A, "solidFill")) {
                            if (!TryGetSolidLine(properties, workbookPart, out _)) {
                                return false;
                            }

                            continue;
                        }

                        return false;
                    }

                    continue;
                }

                return false;
            }

            return true;
        }

        private static bool IsSimpleSupportedDataPoint(OfficeXmlNode point, ExcelPartRef workbookPart) {
            OfficeXmlNode properties = point.Child(C, "spPr");
            return properties.IsNone || ImageExportIsSupportedPointStyle(properties, ImageExportColorScheme(workbookPart));
        }

        private static string? GetAxisTitleText(OfficeXmlNode title) {
            if (title.IsNone) {
                return null;
            }

            string text = string.Concat(title.Descendants(A, "t").Select(item => item.InnerText));
            return string.IsNullOrWhiteSpace(text) ? null : text.Trim();
        }

        private static OfficeChartLegendPosition MapLegendPosition(string? position) => position switch {
            null => OfficeChartLegendPosition.Right,
            "l" => OfficeChartLegendPosition.Left,
            "t" => OfficeChartLegendPosition.Top,
            "b" => OfficeChartLegendPosition.Bottom,
            "tr" => OfficeChartLegendPosition.TopRight,
            _ => OfficeChartLegendPosition.Right
        };

        private static OfficeChartDataLabelPosition MapDataLabelPosition(string? position) => position switch {
            "ctr" => OfficeChartDataLabelPosition.Center,
            "inBase" => OfficeChartDataLabelPosition.InsideBase,
            "inEnd" => OfficeChartDataLabelPosition.InsideEnd,
            "outEnd" => OfficeChartDataLabelPosition.OutsideEnd,
            "l" => OfficeChartDataLabelPosition.Left,
            "r" => OfficeChartDataLabelPosition.Right,
            "t" => OfficeChartDataLabelPosition.Top,
            "b" => OfficeChartDataLabelPosition.Bottom,
            _ => OfficeChartDataLabelPosition.BestFit
        };

        // The child sequence a series of this chart group allows, or null for a parent that is not a chart group.
        private static OfficeXmlChildOrder? ImageExportSeriesChildOrder(OfficeXmlNode group) {
            if (group.NamespaceId != C) return null;
            switch (group.LocalName) {
                case "barChart": case "bar3DChart": return CBarSerView.Order;
                case "lineChart": case "line3DChart": case "stockChart": return CLineSerView.Order;
                case "areaChart": case "area3DChart": return CAreaSerView.Order;
                case "pieChart": case "pie3DChart": case "doughnutChart": case "ofPieChart": return CPieSerView.Order;
                case "scatterChart": return CScatterSerView.Order;
                case "bubbleChart": return CBubbleSerView.Order;
                case "radarChart": return CRadarSerView.Order;
                case "surfaceChart": case "surface3DChart": return CSurfaceSerView.Order;
                default: return null;
            }
        }

        // Whether an element sits where the chart schema allows it under its series and chart group. Markup a series or
        // chart type does not list, such as data labels on a surface series, is unknown markup to the readers, with
        // everything inside it, so the image export never reads it.
        private static bool IsImageExportSchemaPlaced(OfficeXmlNode element) {
            OfficeXmlNode child = element;
            for (OfficeXmlNode parent = element.Parent; parent.Exists; child = parent, parent = parent.Parent) {
                if (parent.Is(C, "ser")) {
                    OfficeXmlChildOrder? order = ImageExportSeriesChildOrder(parent.Parent);
                    if (order != null && order.RankOf(child.NamespaceId, child.LocalName) < 0) return false;
                } else if (parent.Is(C, "surfaceChart") || parent.Is(C, "surface3DChart")) {
                    if (child.Is(C, "dLbls")) return false;
                } else if (parent.Is(C, "plotArea")) {
                    break;
                }
            }
            return true;
        }

        // The first child of a series with this name when the series type allows it.
        private static OfficeXmlNode ImageExportSeriesChild(OfficeXmlNode series, string localName) {
            OfficeXmlChildOrder? order = ImageExportSeriesChildOrder(series.Parent);
            return order != null && order.RankOf(C, localName) < 0 ? OfficeXmlNode.None : series.Child(C, localName);
        }

        // The series element of any chart group: a c:ser under one of the c: chart types.
        private static bool IsImageExportSeriesElement(OfficeXmlNode element) {
            if (!element.Is(C, "ser")) return false;
            OfficeXmlNode group = element.Parent;
            if (group.NamespaceId != C) return false;
            switch (group.LocalName) {
                case "barChart": case "bar3DChart":
                case "lineChart": case "line3DChart": case "stockChart":
                case "areaChart": case "area3DChart":
                case "pieChart": case "pie3DChart": case "doughnutChart": case "ofPieChart":
                case "scatterChart": case "bubbleChart": case "radarChart":
                case "surfaceChart": case "surface3DChart":
                    return true;
                default:
                    return false;
            }
        }

        // A pie, 3-D pie, doughnut or pie-of-pie series.
        private static bool IsImageExportPieSeries(OfficeXmlNode series) {
            if (!series.Is(C, "ser")) return false;
            OfficeXmlNode group = series.Parent;
            return group.Is(C, "pieChart") || group.Is(C, "pie3DChart") || group.Is(C, "doughnutChart") || group.Is(C, "ofPieChart");
        }

        private static OfficeXmlNode FirstImageExportPieSeries(OfficeXmlNode plotArea) {
            if (plotArea.IsNone) return OfficeXmlNode.None;
            foreach (OfficeXmlNode series in plotArea.Descendants(C, "ser")) {
                if (IsImageExportPieSeries(series)) return series;
            }
            return OfficeXmlNode.None;
        }

        // The theme's colour scheme, as the shared chart readers resolve scheme colours with it.
        private static OfficeXmlNode ImageExportColorScheme(ExcelPartRef workbookPart) =>
            workbookPart.Theme.Root.Child(A, "themeElements").Child(A, "clrScheme");

        // An unsigned byte value, as the schema's xsd:unsignedByte settings hold.
        private static int? ReadImageExportByteVal(OfficeXmlNode element) =>
            element.Int32Val() is int value && value >= byte.MinValue && value <= byte.MaxValue ? value : null;

        private sealed class ImageExportSeriesStyle {
            internal string? SeriesColorArgb { get; set; }

            internal OfficeChartTrendline? Trendline { get; set; }

            internal IReadOnlyList<OfficeChartErrorBars>? ErrorBars { get; set; }

            internal bool AutomaticColor { get; set; }

            internal double? SeriesLineWidth { get; set; }

            internal OfficeStrokeDashStyle? SeriesLineDashStyle { get; set; }

            internal IReadOnlyList<string?>? PointColorArgb { get; set; }

            internal OfficeChartPointStyle? SeriesStyle { get; set; }

            internal IReadOnlyList<OfficeChartPointStyle?>? PointStyles { get; set; }

            internal IReadOnlyList<int>? PointExplosions { get; set; }

            internal bool ShowMarkers { get; set; } = true;

            internal bool? ConnectLine { get; set; }

            internal int? MarkerSize { get; set; }

            internal OfficeChartMarkerShape? MarkerShape { get; set; }

            internal string? MarkerOutlineColorArgb { get; set; }

            internal double? MarkerOutlineWidth { get; set; }

            internal bool HasAny => SeriesStyle != null || PointStyles != null || PointExplosions != null || SeriesColorArgb != null || SeriesLineWidth != null || SeriesLineDashStyle != null || PointColorArgb != null || !ShowMarkers || ConnectLine.HasValue || MarkerSize != null || MarkerShape != null || MarkerOutlineColorArgb != null || MarkerOutlineWidth != null;
        }
    }
}
