using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using OfficeIMO.Drawing;
using OfficeIMO.Excel.Utilities;

namespace OfficeIMO.Excel {
    public sealed partial class ExcelChart {
        private OfficeChartStyle? CreateImageExportStyle() {
            C.ChartSpace? chartSpace = GetChartPart().ChartSpace;
            if (chartSpace == null) {
                return null;
            }

            C.Chart? chart = chartSpace.GetFirstChild<C.Chart>();
            C.PlotArea? plotArea = chart?.GetFirstChild<C.PlotArea>();
            C.ShapeProperties? chartAreaProperties = chartSpace.GetFirstChild<C.ShapeProperties>();
            OpenXmlCompositeElement? plotAreaProperties = plotArea?.GetFirstChild<C.ShapeProperties>();
            WorkbookPart workbookPart = _document.WorkbookPartRoot;

            OfficeColor? chartFill = TryGetSolidFill(chartAreaProperties, workbookPart, out OfficeColor resolvedChartFill) ? resolvedChartFill : null;
            OfficeColor? chartLine = TryGetSolidLine(chartAreaProperties, workbookPart, out OfficeColor resolvedChartLine) ? resolvedChartLine : null;
            OfficeColor? plotFill = TryGetSolidFill(plotAreaProperties, workbookPart, out OfficeColor resolvedPlotFill) ? resolvedPlotFill : null;
            OfficeColor? plotLine = TryGetSolidLine(plotAreaProperties, workbookPart, out OfficeColor resolvedPlotLine) ? resolvedPlotLine : null;
            OpenXmlCompositeElement? categoryAxis = plotArea == null ? null : ResolveImageExportCategoryAxis(plotArea);
            OpenXmlCompositeElement? valueAxis = plotArea == null ? null : ResolveImageExportValueAxis(plotArea);
            double? chartLineWidth = TryGetLineWidth(chartAreaProperties, out double resolvedChartLineWidth) ? resolvedChartLineWidth : null;
            double? plotLineWidth = TryGetLineWidth(plotAreaProperties, out double resolvedPlotLineWidth) ? resolvedPlotLineWidth : null;
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
            string? titleFontFamily = TryGetImageExportTitleFontFamily(chart, out string? resolvedTitleFontFamily) ? resolvedTitleFontFamily : null;
            double? titleFontSize = TryGetImageExportTitleFontSize(chart, out double resolvedTitleFontSize) ? resolvedTitleFontSize : null;
            OfficeFontStyle? titleFontStyle = TryGetImageExportTitleFontStyle(chart, out OfficeFontStyle resolvedTitleFontStyle) ? resolvedTitleFontStyle : null;
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
            bool hasGridLineVisibility = plotArea != null && HasImageExportCartesianAxis(plotArea);
            bool showCategoryGridLines = HasImageExportMajorGridlines(categoryAxis);
            bool showValueGridLines = HasImageExportMajorGridlines(valueAxis);
            bool showCategoryMinorGridLines = HasImageExportMinorGridlines(categoryAxis);
            bool showValueMinorGridLines = HasImageExportMinorGridlines(valueAxis);
            if (chartFill == null &&
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

            var style = new OfficeChartStyle(
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
            C.Chart? chart = GetChartPart().ChartSpace?.GetFirstChild<C.Chart>();
            C.PlotArea? plotArea = chart?.GetFirstChild<C.PlotArea>();
            if (chart == null || plotArea == null) {
                return null;
            }

            C.DataLabels? dataLabels = plotArea.Descendants<C.DataLabels>().FirstOrDefault(HasAnyVisibleDataLabelPart);
            C.Legend? legend = chart.GetFirstChild<C.Legend>();
            C.Title? title = chart.GetFirstChild<C.Title>();
            OpenXmlCompositeElement? categoryAxis = ResolveImageExportCategoryAxis(plotArea);
            OpenXmlCompositeElement? valueAxis = ResolveImageExportValueAxis(plotArea);
            string? categoryAxisTitle = GetAxisTitleText(categoryAxis?.GetFirstChild<C.Title>());
            string? valueAxisTitle = GetAxisTitleText(valueAxis?.GetFirstChild<C.Title>());
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
            OfficeChartAxisCrossingPosition horizontalAxisCrossingPosition = GetImageExportHorizontalAxisCrossingPosition(plotArea, categoryAxis);
            OfficeChartAxisCrossingPosition verticalAxisCrossingPosition = GetImageExportVerticalAxisCrossingPosition(plotArea, valueAxis);
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
            C.ScatterStyleValues? scatterStyle = GetImageExportScatterStyle(plotArea);
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
            bool hasAxisCrossingPosition = horizontalAxisCrossingPosition != OfficeChartAxisCrossingPosition.AutoZero ||
                verticalAxisCrossingPosition != OfficeChartAxisCrossingPosition.AutoZero;
            bool hasCategoryAxisOrientation = categoryAxisOrientationSpecified;
            bool fillRadarSeries = GetImageExportFillRadarSeries(plotArea);
            bool hasRadarFillLayout = !fillRadarSeries;
            bool hasScatterStyleLayout = !connectScatterPoints;
            bool hasLegendLayout = legend == null ||
                legend.GetFirstChild<C.LegendPosition>() != null ||
                legend.GetFirstChild<C.Overlay>() != null;
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
                dataLabels != null ||
                hasLegendLayout ||
                title?.GetFirstChild<C.Overlay>() != null ||
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
                hasAxisCrossingPosition ||
                hasCategoryAxisOrientation ||
                hasRadarFillLayout ||
                hasScatterStyleLayout ||
                hasTextFont;
            if (!hasLayout) {
                return null;
            }

            bool showLegend = legend != null;
            OfficeChartLegendPosition legendPosition = MapLegendPosition(legend?.GetFirstChild<C.LegendPosition>()?.Val?.Value);
            bool overlayLegend = IsEnabled(legend?.GetFirstChild<C.Overlay>());
            bool showDataLabels = dataLabels != null;
            return new OfficeChartLayout(
                overlayLegend: overlayLegend,
                showLegend: showLegend,
                legendPosition: legendPosition,
                showDataLabels: showDataLabels,
                showDataLabelValues: IsEnabled(dataLabels?.GetFirstChild<C.ShowValue>()),
                showDataLabelPercentages: IsEnabled(dataLabels?.GetFirstChild<C.ShowPercent>()),
                showDataLabelCategoryNames: IsEnabled(dataLabels?.GetFirstChild<C.ShowCategoryName>()),
                showDataLabelSeriesNames: IsEnabled(dataLabels?.GetFirstChild<C.ShowSeriesName>()),
                dataLabelSeparator: dataLabels?.GetFirstChild<C.Separator>()?.Text,
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
                dataLabelPosition: MapDataLabelPosition(dataLabels?.GetFirstChild<C.DataLabelPosition>()?.Val?.Value),
                dataLabelNumberFormat: dataLabels?.GetFirstChild<C.NumberingFormat>()?.FormatCode?.Value,
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
                horizontalAxisCrossingPosition: horizontalAxisCrossingPosition,
                verticalAxisCrossingPosition: verticalAxisCrossingPosition,
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
                overlayTitle: IsEnabled(title?.GetFirstChild<C.Overlay>()));
        }

        private static ExcelChartData ApplyImageExportSeriesStyles(ChartPart chartPart, ExcelChartData data, WorkbookPart workbookPart) {
            C.PlotArea? plotArea = chartPart.ChartSpace?
                .GetFirstChild<C.Chart>()?
                .GetFirstChild<C.PlotArea>();
            if (plotArea == null || data.Series.Count == 0) {
                return data;
            }

            Dictionary<int, ImageExportSeriesStyle> styles = GetImageExportSeriesStyles(plotArea, data, workbookPart);
            if (styles.Count == 0) {
                return data;
            }

            var series = new List<ExcelChartSeries>(data.Series.Count);
            bool changed = false;
            for (int i = 0; i < data.Series.Count; i++) {
                ExcelChartSeries current = data.Series[i];
                if (styles.TryGetValue(i, out ImageExportSeriesStyle? style)) {
                    series.Add(current.WithImageExportStyle(style.SeriesColorArgb, style.SeriesLineWidth, style.SeriesLineDashStyle, style.PointColorArgb, style.ShowMarkers, style.ConnectLine, style.MarkerSize, style.MarkerShape, style.MarkerOutlineColorArgb, style.MarkerOutlineWidth));
                    changed = true;
                } else {
                    series.Add(current);
                }
            }

            return changed ? new ExcelChartData(data.Categories, series) : data;
        }

        private static Dictionary<int, ImageExportSeriesStyle> GetImageExportSeriesStyles(C.PlotArea plotArea, ExcelChartData data, WorkbookPart workbookPart) {
            var styles = new Dictionary<int, ImageExportSeriesStyle>();
            int seriesOrder = 0;
            foreach (OpenXmlCompositeElement series in plotArea.Descendants<OpenXmlCompositeElement>()
                         .Where(IsSeriesElement)
                         .OrderBy(item => item.GetFirstChild<C.Index>()?.Val?.Value ?? uint.MaxValue)) {
                int index = seriesOrder++;
                if (index < 0 || index >= data.Series.Count) {
                    continue;
                }

                ImageExportSeriesStyle style = new ImageExportSeriesStyle();

                C.ChartShapeProperties? properties = series.GetFirstChild<C.ChartShapeProperties>();
                if (properties != null) {
                    style.SeriesColorArgb = GetImageExportSeriesColor(properties, workbookPart);
                    style.SeriesLineWidth = GetImageExportSeriesLineWidth(properties);
                    style.SeriesLineDashStyle = GetImageExportSeriesLineDashStyle(properties);
                    if (HasNoLine(properties)) {
                        style.ConnectLine = false;
                    }
                }

                int valueCount = data.Series[index].Values.Count;
                C.Marker? marker = series.GetFirstChild<C.Marker>();
                C.ScatterStyleValues? scatterStyle = (series.Parent as C.ScatterChart)?.GetFirstChild<C.ScatterStyle>()?.Val?.Value;
                style.ShowMarkers = GetImageExportShowMarkers(marker, scatterStyle);
                style.MarkerSize = GetImageExportMarkerSize(marker);
                style.MarkerShape = GetImageExportMarkerShape(marker);
                style.MarkerOutlineColorArgb = GetImageExportMarkerOutlineColor(marker, workbookPart);
                style.MarkerOutlineWidth = GetImageExportMarkerOutlineWidth(marker);
                string? markerFill = GetImageExportMarkerFillColor(marker, workbookPart);
                style.PointColorArgb = GetImageExportPointColors(series, valueCount, markerFill, workbookPart);

                if (style.HasAny && !styles.ContainsKey(index)) {
                    styles.Add(index, style);
                }
            }

            return styles;
        }

        private static bool GetImageExportFillRadarSeries(C.PlotArea plotArea) {
            C.RadarChart? radarChart = plotArea.GetFirstChild<C.RadarChart>();
            C.RadarStyleValues? radarStyle = radarChart?.GetFirstChild<C.RadarStyle>()?.Val?.Value;
            return radarStyle == null || radarStyle.Value != C.RadarStyleValues.Standard;
        }

        private static C.ScatterStyleValues? GetImageExportScatterStyle(C.PlotArea plotArea) =>
            plotArea.GetFirstChild<C.ScatterChart>()?.GetFirstChild<C.ScatterStyle>()?.Val?.Value;

        private static bool GetImageExportConnectScatterPoints(C.ScatterStyleValues? style) =>
            style == null || style.Value != C.ScatterStyleValues.Marker;

        private static string? GetImageExportSeriesColor(C.ChartShapeProperties properties, WorkbookPart workbookPart) {
            if (TryGetSolidFill(properties, workbookPart, out OfficeColor fill)) {
                return fill.ToRgbHex();
            }

            if (TryGetSolidLine(properties, workbookPart, out OfficeColor line)) {
                return line.ToRgbHex();
            }

            return null;
        }

        private static double? GetImageExportSeriesLineWidth(C.ChartShapeProperties properties) =>
            TryGetLineWidth(properties, out double width) ? width : null;

        private static OfficeStrokeDashStyle? GetImageExportSeriesLineDashStyle(C.ChartShapeProperties properties) =>
            TryGetLineDashStyle(properties, out OfficeStrokeDashStyle dashStyle) ? dashStyle : null;

        private static bool TryGetImageExportAxisLineColor(OpenXmlCompositeElement? axis, WorkbookPart workbookPart, out OfficeColor color) {
            color = default;
            if (axis == null) {
                return false;
            }

            C.ShapeProperties? properties = axis.GetFirstChild<C.ShapeProperties>();
            return TryGetSolidLine(properties, workbookPart, out color);
        }

        private static bool TryGetImageExportAxisLineWidth(OpenXmlCompositeElement? axis, out double width) {
            width = default;
            if (axis == null) {
                return false;
            }

            C.ShapeProperties? properties = axis.GetFirstChild<C.ShapeProperties>();
            return TryGetLineWidth(properties, out width);
        }

        private static bool TryGetImageExportAxisLineDashStyle(OpenXmlCompositeElement? axis, out OfficeStrokeDashStyle dashStyle) {
            dashStyle = default;
            if (axis == null) {
                return false;
            }

            C.ShapeProperties? properties = axis.GetFirstChild<C.ShapeProperties>();
            return TryGetLineDashStyle(properties, out dashStyle);
        }

        private static bool TryGetImageExportGridLineColor(OpenXmlCompositeElement? axis, WorkbookPart workbookPart, out OfficeColor color) {
            color = default;
            return TryGetImageExportGridLineColorCore(axis?.GetFirstChild<C.MajorGridlines>(), workbookPart, out color);
        }

        private static bool TryGetImageExportMinorGridLineColor(OpenXmlCompositeElement? axis, WorkbookPart workbookPart, out OfficeColor color) {
            color = default;
            return TryGetImageExportGridLineColorCore(axis?.GetFirstChild<C.MinorGridlines>(), workbookPart, out color);
        }

        private static bool TryGetImageExportGridLineColorCore(OpenXmlCompositeElement? gridlines, WorkbookPart workbookPart, out OfficeColor color) {
            color = default;
            C.ChartShapeProperties? properties = gridlines?.GetFirstChild<C.ChartShapeProperties>();
            return TryGetSolidLine(properties, workbookPart, out color);
        }

        private static bool TryGetImageExportGridLineWidth(OpenXmlCompositeElement? axis, out double width) {
            width = default;
            return TryGetImageExportGridLineWidthCore(axis?.GetFirstChild<C.MajorGridlines>(), out width);
        }

        private static bool TryGetImageExportMinorGridLineWidth(OpenXmlCompositeElement? axis, out double width) {
            width = default;
            return TryGetImageExportGridLineWidthCore(axis?.GetFirstChild<C.MinorGridlines>(), out width);
        }

        private static bool TryGetImageExportGridLineWidthCore(OpenXmlCompositeElement? gridlines, out double width) {
            width = default;
            C.ChartShapeProperties? properties = gridlines?.GetFirstChild<C.ChartShapeProperties>();
            return TryGetLineWidth(properties, out width);
        }

        private static bool TryGetImageExportGridLineDashStyle(OpenXmlCompositeElement? axis, out OfficeStrokeDashStyle dashStyle) {
            dashStyle = default;
            return TryGetImageExportGridLineDashStyleCore(axis?.GetFirstChild<C.MajorGridlines>(), out dashStyle);
        }

        private static bool TryGetImageExportMinorGridLineDashStyle(OpenXmlCompositeElement? axis, out OfficeStrokeDashStyle dashStyle) {
            dashStyle = default;
            return TryGetImageExportGridLineDashStyleCore(axis?.GetFirstChild<C.MinorGridlines>(), out dashStyle);
        }

        private static bool TryGetImageExportGridLineDashStyleCore(OpenXmlCompositeElement? gridlines, out OfficeStrokeDashStyle dashStyle) {
            dashStyle = default;
            C.ChartShapeProperties? properties = gridlines?.GetFirstChild<C.ChartShapeProperties>();
            return TryGetLineDashStyle(properties, out dashStyle);
        }

        private static bool HasImageExportCartesianAxis(C.PlotArea plotArea) =>
            GetImageExportAxes(plotArea).Any();

        private static bool HasImageExportMajorGridlines(C.PlotArea plotArea) =>
            GetImageExportMajorGridlines(plotArea).Any();

        private static bool HasImageExportMajorGridlines(OpenXmlCompositeElement? axis) =>
            axis?.GetFirstChild<C.MajorGridlines>() != null;

        private static bool HasImageExportMinorGridlines(OpenXmlCompositeElement? axis) =>
            axis?.GetFirstChild<C.MinorGridlines>() != null;

        private static IEnumerable<OpenXmlCompositeElement> GetImageExportAxes(C.PlotArea plotArea) =>
            plotArea.Elements<C.CategoryAxis>().Cast<OpenXmlCompositeElement>()
                .Concat(plotArea.Elements<C.ValueAxis>())
                .Concat(plotArea.Elements<C.DateAxis>());

        private static IEnumerable<C.MajorGridlines> GetImageExportMajorGridlines(C.PlotArea plotArea) =>
            GetImageExportAxes(plotArea)
                .Select(axis => axis.GetFirstChild<C.MajorGridlines>())
                .OfType<C.MajorGridlines>();

        private static OpenXmlCompositeElement? ResolveImageExportCategoryAxis(C.PlotArea plotArea) =>
            (OpenXmlCompositeElement?)ResolveCategoryAxis(plotArea, ExcelChartAxisGroup.Primary) ?? ResolveScatterXAxis(plotArea);

        private static OpenXmlCompositeElement? ResolveImageExportValueAxis(C.PlotArea plotArea) =>
            (OpenXmlCompositeElement?)ResolveValueAxis(plotArea, ExcelChartAxisGroup.Primary) ?? ResolveScatterYAxis(plotArea);

        private static bool IsImageExportAxisVisible(OpenXmlCompositeElement? axis) =>
            axis == null || !IsEnabled(axis.GetFirstChild<C.Delete>());

        private static bool IsImageExportAxisLineVisible(OpenXmlCompositeElement? axis) =>
            axis == null || (IsImageExportAxisVisible(axis) && !HasNoLine(axis.GetFirstChild<C.ShapeProperties>()));

        private static bool IsImageExportAxisLabelsVisible(OpenXmlCompositeElement? axis) =>
            axis == null || (IsImageExportAxisVisible(axis) && axis.GetFirstChild<C.TickLabelPosition>()?.Val?.Value != C.TickLabelPositionValues.None);

        private static IReadOnlyList<string?>? GetImageExportPointColors(OpenXmlCompositeElement series, int valueCount, string? markerFill, WorkbookPart workbookPart) {
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

            foreach (C.DataPoint point in series.Elements<C.DataPoint>()) {
                uint? rawIndex = point.GetFirstChild<C.Index>()?.Val?.Value;
                if (rawIndex == null || rawIndex.Value > int.MaxValue) {
                    continue;
                }

                int index = (int)rawIndex.Value;
                if (index >= colors.Length) {
                    continue;
                }

                C.ChartShapeProperties? properties = point.GetFirstChild<C.ChartShapeProperties>();
                if (properties == null || !TryGetSolidFill(properties, workbookPart, out OfficeColor color)) {
                    continue;
                }

                colors[index] = color.ToRgbHex();
                any = true;
            }

            return any ? colors : null;
        }

        private static bool GetImageExportShowMarkers(C.Marker? marker, C.ScatterStyleValues? scatterStyle) {
            if (marker == null) {
                return scatterStyle.HasValue && GetImageExportScatterStyleShowMarkers(scatterStyle);
            }

            C.MarkerStyleValues? symbol = marker?.Symbol?.Val?.Value;
            return symbol == null || symbol.Value != C.MarkerStyleValues.None;
        }

        private static bool GetImageExportScatterStyleShowMarkers(C.ScatterStyleValues? style) {
            if (style == null) {
                return true;
            }

            return style.Value == C.ScatterStyleValues.Marker ||
                style.Value == C.ScatterStyleValues.LineMarker ||
                style.Value == C.ScatterStyleValues.SmoothMarker;
        }

        private static int? GetImageExportMarkerSize(C.Marker? marker) {
            byte? size = marker?.Size?.Val?.Value;
            return size.HasValue && size.Value > 0 ? size.Value : null;
        }

        private static OfficeChartMarkerShape? GetImageExportMarkerShape(C.Marker? marker) {
            if (marker == null) {
                return null;
            }

            C.MarkerStyleValues? symbol = marker?.Symbol?.Val?.Value;
            if (symbol == null || symbol.Value == C.MarkerStyleValues.Auto || symbol.Value == C.MarkerStyleValues.Circle) {
                return OfficeChartMarkerShape.Circle;
            }

            if (symbol.Value == C.MarkerStyleValues.Square) {
                return OfficeChartMarkerShape.Square;
            }

            if (symbol.Value == C.MarkerStyleValues.Diamond) {
                return OfficeChartMarkerShape.Diamond;
            }

            if (symbol.Value == C.MarkerStyleValues.Triangle) {
                return OfficeChartMarkerShape.Triangle;
            }

            if (symbol.Value == C.MarkerStyleValues.Dash) {
                return OfficeChartMarkerShape.Dash;
            }

            if (symbol.Value == C.MarkerStyleValues.Dot) {
                return OfficeChartMarkerShape.Dot;
            }

            if (symbol.Value == C.MarkerStyleValues.Plus) {
                return OfficeChartMarkerShape.Plus;
            }

            if (symbol.Value == C.MarkerStyleValues.X) {
                return OfficeChartMarkerShape.X;
            }

            if (symbol.Value == C.MarkerStyleValues.Star) {
                return OfficeChartMarkerShape.Star;
            }

            return null;
        }

        private static string? GetImageExportMarkerFillColor(C.Marker? marker, WorkbookPart workbookPart) {
            C.ChartShapeProperties? properties = marker?.GetFirstChild<C.ChartShapeProperties>();
            return properties != null && TryGetSolidFill(properties, workbookPart, out OfficeColor color) ? color.ToRgbHex() : null;
        }

        private static string? GetImageExportMarkerOutlineColor(C.Marker? marker, WorkbookPart workbookPart) {
            C.ChartShapeProperties? properties = marker?.GetFirstChild<C.ChartShapeProperties>();
            return properties != null && TryGetSolidLine(properties, workbookPart, out OfficeColor color) ? color.ToRgbHex() : null;
        }

        private static double? GetImageExportMarkerOutlineWidth(C.Marker? marker) {
            A.Outline? outline = marker?
                .GetFirstChild<C.ChartShapeProperties>()?
                .GetFirstChild<A.Outline>();
            long? emu = outline?.Width?.Value;
            return emu.HasValue && emu.Value > 0
                ? ExcelImageExportLimits.ClampStrokeWidth(emu.Value / 12700D)
                : null;
        }

        private IReadOnlyList<OfficeImageExportDiagnostic> CreateImageExportDiagnostics() {
            C.ChartSpace? chartSpace = GetChartPart().ChartSpace;
            if (chartSpace == null) {
                return Array.Empty<OfficeImageExportDiagnostic>();
            }

            var diagnostics = new List<OfficeImageExportDiagnostic>();
            string source = GetImageExportSource();
            WorkbookPart workbookPart = _document.WorkbookPartRoot;
            if (chartSpace.Descendants<C.Trendline>().Any()) {
                diagnostics.Add(ExcelImageExportDiagnosticClassifier.Create(
                    OfficeImageExportDiagnosticSeverity.Warning,
                    ExcelImageExportDiagnosticCodes.ChartTrendlineUnsupported,
                    "Worksheet chart trendlines are not rendered by the shared image renderer yet.",
                    source));
            }

            if (chartSpace.Descendants<C.DataLabel>().Any()) {
                diagnostics.Add(ExcelImageExportDiagnosticClassifier.Create(
                    OfficeImageExportDiagnosticSeverity.Warning,
                    ExcelImageExportDiagnosticCodes.ChartDataLabelPointOverridesApproximated,
                    "Worksheet chart point-level data-label overrides are approximated as chart-level data labels in image export.",
                    source));
            }

            if (chartSpace.Descendants<C.LeaderLines>().Any()) {
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

            if (chartSpace.Descendants<C.ChartShapeProperties>().Any(properties => IsUnsupportedImageExportSeriesOrMarkerShapeProperties(properties, workbookPart))) {
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

        private static bool HasAnyVisibleDataLabelPart(C.DataLabels labels) =>
            IsEnabled(labels.GetFirstChild<C.ShowValue>()) ||
            IsEnabled(labels.GetFirstChild<C.ShowPercent>()) ||
            IsEnabled(labels.GetFirstChild<C.ShowCategoryName>()) ||
            IsEnabled(labels.GetFirstChild<C.ShowSeriesName>());

        private static bool IsEnabled(C.BooleanType? value) =>
            value != null && (value.Val?.Value ?? true);

        private static bool TryGetSolidFill(OpenXmlCompositeElement? properties, out OfficeColor color) =>
            TryGetSolidFill(properties, null, out color);

        private static bool TryGetSolidFill(OpenXmlCompositeElement? properties, WorkbookPart? workbookPart, out OfficeColor color) {
            string? value = ExcelThemeColorResolver.Resolve(properties?.GetFirstChild<A.SolidFill>(), workbookPart);
            return TryParseResolvedDrawingColor(value, out color);
        }

        private static bool TryGetSolidLine(OpenXmlCompositeElement? properties, out OfficeColor color) =>
            TryGetSolidLine(properties, null, out color);

        private static bool TryGetSolidLine(OpenXmlCompositeElement? properties, WorkbookPart? workbookPart, out OfficeColor color) {
            string? value = ExcelThemeColorResolver.Resolve(
                properties?
                    .GetFirstChild<A.Outline>()?
                    .GetFirstChild<A.SolidFill>(),
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

        private static bool HasNoFill(OpenXmlCompositeElement? properties) =>
            properties?.GetFirstChild<A.NoFill>() != null;

        private static bool HasNoLine(OpenXmlCompositeElement? properties) =>
            properties?.GetFirstChild<A.Outline>()?.GetFirstChild<A.NoFill>() != null;

        private static bool TryGetLineWidth(OpenXmlCompositeElement? properties, out double width) {
            width = default;
            long? emu = properties?
                .GetFirstChild<A.Outline>()?
                .Width?
                .Value;
            if (!emu.HasValue || emu.Value <= 0) {
                return false;
            }

            width = ExcelImageExportLimits.ClampStrokeWidth(emu.Value / 12700D);
            return true;
        }

        private static bool TryGetLineDashStyle(OpenXmlCompositeElement? properties, out OfficeStrokeDashStyle dashStyle) {
            dashStyle = default;
            A.PresetDash? dash = properties?
                .GetFirstChild<A.Outline>()?
                .GetFirstChild<A.PresetDash>();
            A.PresetLineDashValues? value = dash?.Val?.Value;
            if (value == null || value.Value == A.PresetLineDashValues.Solid) {
                return false;
            }

            return TryMapPresetLineDash(value.Value, out dashStyle);
        }

        private static bool TryMapPresetLineDash(A.PresetLineDashValues value, out OfficeStrokeDashStyle dashStyle) {
            return OfficeStrokeDashStyleMapper.TryMapOfficePresetDash(GetPresetLineDashToken(value), out dashStyle);
        }

        private static string? GetPresetLineDashToken(A.PresetLineDashValues value) {
            if (value == A.PresetLineDashValues.Dash) {
                return "dash";
            }

            if (value == A.PresetLineDashValues.LargeDash) {
                return "lgDash";
            }

            if (value == A.PresetLineDashValues.SystemDash) {
                return "sysDash";
            }

            if (value == A.PresetLineDashValues.Dot) {
                return "dot";
            }

            if (value == A.PresetLineDashValues.SystemDot) {
                return "sysDot";
            }

            if (value == A.PresetLineDashValues.DashDot) {
                return "dashDot";
            }

            if (value == A.PresetLineDashValues.LargeDashDot) {
                return "lgDashDot";
            }

            if (value == A.PresetLineDashValues.SystemDashDot) {
                return "sysDashDot";
            }

            if (value == A.PresetLineDashValues.LargeDashDotDot) {
                return "lgDashDotDot";
            }

            if (value == A.PresetLineDashValues.SystemDashDotDot) {
                return "sysDashDotDot";
            }

            return null;
        }

        private static bool HasUnsupportedImageExportGridlineStyle(C.ChartSpace chartSpace, WorkbookPart workbookPart) {
            foreach (C.MajorGridlines gridlines in chartSpace.Descendants<C.MajorGridlines>()) {
                C.ChartShapeProperties? properties = gridlines.GetFirstChild<C.ChartShapeProperties>();
                if (properties != null && !IsSimpleSupportedGridlineShapeProperties(properties, workbookPart)) {
                    return true;
                }
            }

            foreach (C.MinorGridlines gridlines in chartSpace.Descendants<C.MinorGridlines>()) {
                C.ChartShapeProperties? properties = gridlines.GetFirstChild<C.ChartShapeProperties>();
                if (properties != null && !IsSimpleSupportedGridlineShapeProperties(properties, workbookPart)) {
                    return true;
                }
            }

            return false;
        }

        private static bool IsSimpleSupportedGridlineShapeProperties(C.ChartShapeProperties properties, WorkbookPart workbookPart) {
            if (!properties.ChildElements.Any()) {
                return true;
            }

            foreach (OpenXmlElement child in properties.ChildElements) {
                if (child is A.Outline outline) {
                    foreach (OpenXmlElement outlineChild in outline.ChildElements) {
                        if (outlineChild is A.SolidFill) {
                            if (!TryGetSolidLine(properties, workbookPart, out _)) {
                                return false;
                            }

                            continue;
                        }

                        if (outlineChild is A.PresetDash dash) {
                            A.PresetLineDashValues? value = dash.Val?.Value;
                            if (value != null && value.Value != A.PresetLineDashValues.Solid && !TryMapPresetLineDash(value.Value, out _)) {
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

        private static bool HasUnsupportedImageExportAxisStyle(C.ChartSpace chartSpace, WorkbookPart workbookPart) {
            foreach (C.PlotArea plotArea in chartSpace.Descendants<C.PlotArea>()) {
                foreach (OpenXmlCompositeElement axis in GetImageExportAxes(plotArea)) {
                    C.ShapeProperties? properties = axis.GetFirstChild<C.ShapeProperties>();
                    if (properties != null && !IsSimpleSupportedAxisShapeProperties(properties, workbookPart)) {
                        return true;
                    }
                }
            }

            return false;
        }

        private static bool IsSimpleSupportedAxisShapeProperties(C.ShapeProperties properties, WorkbookPart workbookPart) {
            if (!properties.ChildElements.Any()) {
                return true;
            }

            foreach (OpenXmlElement child in properties.ChildElements) {
                if (child is A.Outline outline) {
                    foreach (OpenXmlElement outlineChild in outline.ChildElements) {
                        if (outlineChild is A.SolidFill) {
                            if (!TryGetSolidLine(properties, workbookPart, out _)) {
                                return false;
                            }

                            continue;
                        }

                        if (outlineChild is A.NoFill) {
                            continue;
                        }

                        if (outlineChild is A.PresetDash dash) {
                            A.PresetLineDashValues? value = dash.Val?.Value;
                            if (value != null && value.Value != A.PresetLineDashValues.Solid && !TryMapPresetLineDash(value.Value, out _)) {
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

        private static bool HasUnsupportedImageExportAxisTickLabelPosition(C.ChartSpace chartSpace) {
            foreach (C.PlotArea plotArea in chartSpace.Descendants<C.PlotArea>()) {
                foreach (OpenXmlCompositeElement axis in GetImageExportAxes(plotArea)) {
                    C.TickLabelPositionValues? position = axis.GetFirstChild<C.TickLabelPosition>()?.Val?.Value;
                    if (position != null &&
                        position.Value != C.TickLabelPositionValues.NextTo &&
                        position.Value != C.TickLabelPositionValues.None &&
                        position.Value != C.TickLabelPositionValues.High &&
                        position.Value != C.TickLabelPositionValues.Low) {
                        return true;
                    }
                }
            }

            return false;
        }

        private static bool HasApproximatedImageExportAxisMinorTickMarks(C.ChartSpace chartSpace) {
            foreach (C.PlotArea plotArea in chartSpace.Descendants<C.PlotArea>()) {
                foreach (OpenXmlCompositeElement axis in GetImageExportAxes(plotArea)) {
                    if (HasApproximatedImageExportAxisMinorTickMark(axis)) {
                        return true;
                    }
                }
            }

            return false;
        }

        private static bool HasApproximatedImageExportAxisMinorTickMark(OpenXmlCompositeElement axis) {
            C.TickMarkValues? tickMark = axis.GetFirstChild<C.MinorTickMark>()?.Val?.Value;
            if (tickMark == null || tickMark.Value == C.TickMarkValues.None) {
                return false;
            }

            if (axis is not C.ValueAxis) {
                return true;
            }

            double? minorUnit = axis.GetFirstChild<C.MinorUnit>()?.Val?.Value;
            return !minorUnit.HasValue ||
                   double.IsNaN(minorUnit.Value) ||
                   double.IsInfinity(minorUnit.Value) ||
                   minorUnit.Value <= 0D;
        }

        private static bool HasUnsupportedImageExportAxisCrossing(C.ChartSpace chartSpace) {
            foreach (C.PlotArea plotArea in chartSpace.Descendants<C.PlotArea>()) {
                foreach (OpenXmlCompositeElement axis in GetImageExportAxes(plotArea)) {
                    if (axis.GetFirstChild<C.CrossesAt>() != null) {
                        return true;
                    }

                    C.CrossesValues? crosses = axis.GetFirstChild<C.Crosses>()?.Val?.Value;
                    if (crosses != null &&
                        crosses.Value != C.CrossesValues.AutoZero &&
                        !IsSupportedImageExportAxisCrossing(plotArea, axis, crosses.Value)) {
                        return true;
                    }
                }
            }

            return false;
        }

        private static bool IsSupportedImageExportAxisCrossing(C.PlotArea plotArea, OpenXmlCompositeElement axis, C.CrossesValues crosses) =>
            crosses == C.CrossesValues.Maximum &&
            (axis is C.ValueAxis || axis is C.CategoryAxis || axis is C.DateAxis) &&
            !HasHorizontalBarChart(plotArea);

        private static bool HasUnsupportedImageExportAxisScale(C.ChartSpace chartSpace) {
            foreach (C.PlotArea plotArea in chartSpace.Descendants<C.PlotArea>()) {
                foreach (OpenXmlCompositeElement axis in GetImageExportAxes(plotArea)) {
                    C.Scaling? scaling = axis.GetFirstChild<C.Scaling>();
                    if (scaling != null && HasUnsupportedImageExportAxisScaling(plotArea, axis, scaling)) {
                        return true;
                    }

                    if (axis.GetFirstChild<C.MajorUnit>() != null && axis is not C.ValueAxis) {
                        return true;
                    }

                    C.MinorUnit? minorUnit = axis.GetFirstChild<C.MinorUnit>();
                    double? minorUnitValue = minorUnit?.Val?.Value;
                    if (minorUnit != null &&
                        (axis is not C.ValueAxis ||
                         !minorUnitValue.HasValue ||
                         double.IsNaN(minorUnitValue.Value) ||
                         double.IsInfinity(minorUnitValue.Value) ||
                         minorUnitValue.Value <= 0D)) {
                        return true;
                    }

                    C.CrossBetweenValues? crossBetween = axis.GetFirstChild<C.CrossBetween>()?.Val?.Value;
                    if (crossBetween != null && crossBetween.Value != C.CrossBetweenValues.Between) {
                        return true;
                    }
                }
            }

            return false;
        }

        private static bool HasUnsupportedImageExportAxisScaling(C.PlotArea plotArea, OpenXmlCompositeElement axis, C.Scaling scaling) {
            C.OrientationValues? orientation = scaling.GetFirstChild<C.Orientation>()?.Val?.Value;
            if (orientation != null &&
                orientation.Value != C.OrientationValues.MinMax &&
                !IsSupportedImageExportAxisOrientation(plotArea, axis, orientation.Value)) {
                return true;
            }

            if (scaling.GetFirstChild<C.LogBase>() != null) {
                return true;
            }

            C.MinAxisValue? minimum = scaling.GetFirstChild<C.MinAxisValue>();
            C.MaxAxisValue? maximum = scaling.GetFirstChild<C.MaxAxisValue>();
            if ((minimum != null || maximum != null) && axis is not C.ValueAxis) {
                return true;
            }

            double? min = minimum?.Val?.Value;
            double? max = maximum?.Val?.Value;
            if (min.HasValue && (double.IsNaN(min.Value) || double.IsInfinity(min.Value))) {
                return true;
            }

            if (max.HasValue && (double.IsNaN(max.Value) || double.IsInfinity(max.Value))) {
                return true;
            }

            return min.HasValue && max.HasValue && max.Value <= min.Value;
        }

        private static bool IsSupportedImageExportAxisOrientation(C.PlotArea plotArea, OpenXmlCompositeElement axis, C.OrientationValues orientation) =>
            orientation == C.OrientationValues.MaxMin &&
            (axis is C.CategoryAxis || axis is C.DateAxis) &&
            !HasHorizontalBarChart(plotArea);

        private static bool HasUnsupportedImageExportChartAreaStyle(C.ChartSpace chartSpace, WorkbookPart workbookPart) {
            C.ShapeProperties? chartAreaProperties = chartSpace.GetFirstChild<C.ShapeProperties>();
            if (chartAreaProperties != null && !IsSimpleSupportedChartAreaShapeProperties(chartAreaProperties, workbookPart)) {
                return true;
            }

            foreach (C.PlotArea plotArea in chartSpace.Descendants<C.PlotArea>()) {
                C.ShapeProperties? properties = plotArea.GetFirstChild<C.ShapeProperties>();
                if (properties != null && !IsSimpleSupportedChartAreaShapeProperties(properties, workbookPart)) {
                    return true;
                }
            }

            return false;
        }

        private static bool IsSimpleSupportedChartAreaShapeProperties(C.ShapeProperties properties, WorkbookPart workbookPart) {
            if (!properties.ChildElements.Any()) {
                return true;
            }

            foreach (OpenXmlElement child in properties.ChildElements) {
                if (child is A.SolidFill) {
                    if (!TryGetSolidFill(properties, workbookPart, out _)) {
                        return false;
                    }

                    continue;
                }

                if (child is A.NoFill) {
                    continue;
                }

                if (child is A.Outline outline) {
                    if (outline.Width != null && !TryGetLineWidth(properties, out _)) {
                        return false;
                    }

                    foreach (OpenXmlElement outlineChild in outline.ChildElements) {
                        if (outlineChild is A.SolidFill) {
                            if (!TryGetSolidLine(properties, workbookPart, out _)) {
                                return false;
                            }

                            continue;
                        }

                        if (outlineChild is A.NoFill) {
                            continue;
                        }

                        if (outlineChild is A.PresetDash dash) {
                            A.PresetLineDashValues? value = dash.Val?.Value;
                            if (value != null && value.Value != A.PresetLineDashValues.Solid && !TryMapPresetLineDash(value.Value, out _)) {
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

        private static bool IsSeriesOrMarkerShapeProperties(C.ChartShapeProperties properties) {
            OpenXmlElement? parent = properties.Parent;
            return parent is C.AreaChartSeries
                || parent is C.BarChartSeries
                || parent is C.BubbleChartSeries
                || parent is C.LineChartSeries
                || parent is C.PieChartSeries
                || parent is C.RadarChartSeries
                || parent is C.ScatterChartSeries
                || parent is C.SurfaceChartSeries
                || parent is C.DataPoint
                || parent is C.Marker
                || parent is C.DataLabels
                || parent is C.DataLabel
                || parent is C.Trendline;
        }

        private static bool IsUnsupportedImageExportSeriesOrMarkerShapeProperties(C.ChartShapeProperties properties, WorkbookPart workbookPart) {
            if (!IsSeriesOrMarkerShapeProperties(properties)) {
                return false;
            }

            OpenXmlElement? parent = properties.Parent;
            if (parent is C.Marker marker) {
                return !IsSimpleSupportedMarker(marker, workbookPart);
            }

            if (parent is C.DataPoint point) {
                return !IsSimpleSupportedDataPoint(point, workbookPart);
            }

            if (parent is C.DataLabels) {
                return !IsSimpleSupportedSeriesShapeProperties(properties, workbookPart);
            }

            if (parent is C.DataLabel || parent is C.Trendline) {
                return true;
            }

            return !IsSimpleSupportedSeriesShapeProperties(properties, workbookPart);
        }

        private static bool IsSimpleSupportedSeriesShapeProperties(C.ChartShapeProperties properties, WorkbookPart workbookPart) {
            if (!properties.ChildElements.Any()) {
                return true;
            }

            foreach (OpenXmlElement child in properties.ChildElements) {
                if (child is A.SolidFill) {
                    if (!TryGetSolidFill(properties, workbookPart, out _)) {
                        return false;
                    }

                    continue;
                }

                if (child is A.Outline outline) {
                    if (outline.Width != null && !TryGetLineWidth(properties, out _)) {
                        return false;
                    }

                    foreach (OpenXmlElement outlineChild in outline.ChildElements) {
                        if (outlineChild is A.SolidFill) {
                            if (!TryGetSolidLine(properties, workbookPart, out _)) {
                                return false;
                            }

                            continue;
                        }

                        if (outlineChild is A.PresetDash dash) {
                            A.PresetLineDashValues? value = dash.Val?.Value;
                            if (value != null && value.Value != A.PresetLineDashValues.Solid && !TryMapPresetLineDash(value.Value, out _)) {
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

        private static bool IsSimpleSupportedMarker(C.Marker marker, WorkbookPart workbookPart) {
            C.MarkerStyleValues? symbol = marker.Symbol?.Val?.Value;
            if (symbol != null &&
                symbol.Value != C.MarkerStyleValues.Circle &&
                symbol.Value != C.MarkerStyleValues.Auto &&
                symbol.Value != C.MarkerStyleValues.Square &&
                symbol.Value != C.MarkerStyleValues.Diamond &&
                symbol.Value != C.MarkerStyleValues.Triangle &&
                symbol.Value != C.MarkerStyleValues.Dash &&
                symbol.Value != C.MarkerStyleValues.Dot &&
                symbol.Value != C.MarkerStyleValues.Plus &&
                symbol.Value != C.MarkerStyleValues.X &&
                symbol.Value != C.MarkerStyleValues.Star &&
                symbol.Value != C.MarkerStyleValues.None) {
                return false;
            }

            C.ChartShapeProperties? properties = marker.GetFirstChild<C.ChartShapeProperties>();
            return properties == null || IsSimpleSupportedMarkerShapeProperties(properties, workbookPart);
        }

        private static bool IsSimpleSupportedMarkerShapeProperties(C.ChartShapeProperties properties, WorkbookPart workbookPart) {
            if (!properties.ChildElements.Any()) {
                return true;
            }

            foreach (OpenXmlElement child in properties.ChildElements) {
                if (child is A.SolidFill) {
                    if (!TryGetSolidFill(properties, workbookPart, out _)) {
                        return false;
                    }

                    continue;
                }

                if (child is A.Outline outline) {
                    foreach (OpenXmlElement outlineChild in outline.ChildElements) {
                        if (outlineChild is A.SolidFill) {
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

        private static bool IsSimpleSupportedDataPoint(C.DataPoint point, WorkbookPart workbookPart) {
            C.ChartShapeProperties? properties = point.GetFirstChild<C.ChartShapeProperties>();
            if (properties == null) {
                return true;
            }

            if (!properties.ChildElements.Any()) {
                return true;
            }

            foreach (OpenXmlElement child in properties.ChildElements) {
                if (child is A.SolidFill) {
                    if (!TryGetSolidFill(properties, workbookPart, out _)) {
                        return false;
                    }

                    continue;
                }

                return false;
            }

            return true;
        }

        private static string? GetAxisTitleText(C.Title? title) {
            if (title == null) {
                return null;
            }

            string text = string.Concat(title.Descendants<A.Text>().Select(item => item.Text));
            return string.IsNullOrWhiteSpace(text) ? null : text.Trim();
        }

        private static OfficeChartLegendPosition MapLegendPosition(C.LegendPositionValues? position) {
            if (position == null) {
                return OfficeChartLegendPosition.Right;
            }

            C.LegendPositionValues value = position.Value;
            if (value == C.LegendPositionValues.Left) {
                return OfficeChartLegendPosition.Left;
            }

            if (value == C.LegendPositionValues.Top) {
                return OfficeChartLegendPosition.Top;
            }

            if (value == C.LegendPositionValues.Bottom) {
                return OfficeChartLegendPosition.Bottom;
            }

            return OfficeChartLegendPosition.Right;
        }

        private static OfficeChartDataLabelPosition MapDataLabelPosition(C.DataLabelPositionValues? position) {
            if (position == null) {
                return OfficeChartDataLabelPosition.Auto;
            }

            C.DataLabelPositionValues value = position.Value;
            if (value == C.DataLabelPositionValues.Center) {
                return OfficeChartDataLabelPosition.Center;
            }

            if (value == C.DataLabelPositionValues.InsideBase) {
                return OfficeChartDataLabelPosition.InsideBase;
            }

            if (value == C.DataLabelPositionValues.InsideEnd) {
                return OfficeChartDataLabelPosition.InsideEnd;
            }

            if (value == C.DataLabelPositionValues.OutsideEnd) {
                return OfficeChartDataLabelPosition.OutsideEnd;
            }

            if (value == C.DataLabelPositionValues.Left) {
                return OfficeChartDataLabelPosition.Left;
            }

            if (value == C.DataLabelPositionValues.Right) {
                return OfficeChartDataLabelPosition.Right;
            }

            if (value == C.DataLabelPositionValues.Top) {
                return OfficeChartDataLabelPosition.Top;
            }

            if (value == C.DataLabelPositionValues.Bottom) {
                return OfficeChartDataLabelPosition.Bottom;
            }

            return OfficeChartDataLabelPosition.Auto;
        }

        private sealed class ImageExportSeriesStyle {
            internal string? SeriesColorArgb { get; set; }

            internal double? SeriesLineWidth { get; set; }

            internal OfficeStrokeDashStyle? SeriesLineDashStyle { get; set; }

            internal IReadOnlyList<string?>? PointColorArgb { get; set; }

            internal bool ShowMarkers { get; set; } = true;

            internal bool? ConnectLine { get; set; }

            internal int? MarkerSize { get; set; }

            internal OfficeChartMarkerShape? MarkerShape { get; set; }

            internal string? MarkerOutlineColorArgb { get; set; }

            internal double? MarkerOutlineWidth { get; set; }

            internal bool HasAny => SeriesColorArgb != null || SeriesLineWidth != null || SeriesLineDashStyle != null || PointColorArgb != null || !ShowMarkers || ConnectLine.HasValue || MarkerSize != null || MarkerShape != null || MarkerOutlineColorArgb != null || MarkerOutlineWidth != null;
        }
    }
}
