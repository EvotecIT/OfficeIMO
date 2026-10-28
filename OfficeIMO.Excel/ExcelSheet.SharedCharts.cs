using System;
using System.Collections.Generic;
using System.Linq;
using OfficeIMO.Drawing;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        /// <summary>
        /// Adds a native Excel chart from the shared OfficeIMO chart contract, including per-series combo kinds
        /// and primary or secondary axis assignments.
        /// </summary>
        public ExcelChart AddChart(OfficeChartType chartKind, OfficeChartData data, int row, int column,
            OfficeLength? width = null, OfficeLength? height = null, string? title = null) {
            if (data == null) throw new ArgumentNullException(nameof(data));
            ExcelChartData excelData = ToExcelChartData(data, chartKind);
            ExcelChart chart = AddChart(excelData, row, column, width, height,
                chartKind, title);
            chart.ApplyAuthoredSeriesStyles(excelData.Series,
                data.Series.Select(series => series.ShowInLegend).ToList());
            chart.ApplySharedPointStyles(data.Series);
            chart.ApplySharedPointExplosions(data.Series);
            return chart;
        }

        private static ExcelChartData ToExcelChartData(OfficeChartData data, OfficeChartType defaultKind) {
            var series = new List<ExcelChartSeries>(data.Series.Count);
            foreach (OfficeSeries item in data.Series) {
                OfficeChartType? chartType = item.RenderKind ?? defaultKind;
                if (item.PointExplosions != null &&
                    chartType is not OfficeChartType.Pie and not OfficeChartType.Doughnut)
                    throw new NotSupportedException("Point explosions require a pie or doughnut chart.");
                OfficeChartAxisGroup axisGroup = item.AxisGroup == OfficeChartAxisGroup.Secondary
                    ? OfficeChartAxisGroup.Secondary
                    : OfficeChartAxisGroup.Primary;
                ExcelChartSeries mapped = item.XValues == null
                    ? new ExcelChartSeries(item.Name, item.Values, chartType, axisGroup,
                        item.Color?.ToRgbHex())
                    : new ExcelChartSeries(item.Name, item.Values, item.XValues, chartType, axisGroup,
                        item.Color?.ToRgbHex());
                mapped = mapped.WithImageExportStyle(
                    item.Color?.ToRgbHex(), item.StrokeWidth, item.StrokeDashStyle,
                    item.PointColors?.Select(color => color?.ToRgbHex()).ToList(), item.ShowMarkers,
                    item.ConnectLine, item.MarkerSize, item.MarkerShape, item.MarkerOutlineColor?.ToRgbHex(),
                    item.MarkerOutlineWidth).WithSeriesStyle(item.SeriesStyle)
                    .WithPointStyles(item.PointStyles).WithPointExplosions(item.PointExplosions);
                series.Add(mapped);
            }
            return new ExcelChartData(data.Categories, series);
        }

    }
}
