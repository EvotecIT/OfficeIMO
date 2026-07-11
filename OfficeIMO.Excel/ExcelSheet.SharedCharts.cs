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
        public ExcelChart AddChart(OfficeChartKind chartKind, OfficeChartData data, int row, int column,
            int widthPixels = 640, int heightPixels = 360, string? title = null) {
            if (data == null) throw new ArgumentNullException(nameof(data));
            ExcelChartData excelData = ToExcelChartData(data, chartKind);
            ExcelChart chart = AddChart(excelData, row, column, widthPixels, heightPixels,
                MapChartKind(chartKind), title);
            chart.ApplyAuthoredSeriesStyles(excelData.Series);
            return chart;
        }

        private static ExcelChartData ToExcelChartData(OfficeChartData data, OfficeChartKind defaultKind) {
            var series = new List<ExcelChartSeries>(data.Series.Count);
            foreach (OfficeChartSeries item in data.Series) {
                ExcelChartType? chartType = MapChartKind(item.RenderKind ?? defaultKind);
                ExcelChartAxisGroup axisGroup = item.AxisGroup == OfficeChartAxisGroup.Secondary
                    ? ExcelChartAxisGroup.Secondary
                    : ExcelChartAxisGroup.Primary;
                ExcelChartSeries mapped = item.XValues == null
                    ? new ExcelChartSeries(item.Name, item.Values, chartType, axisGroup,
                        item.Color?.ToRgbHex())
                    : new ExcelChartSeries(item.Name, item.Values, item.XValues, chartType, axisGroup,
                        item.Color?.ToRgbHex());
                mapped = mapped.WithImageExportStyle(
                    item.Color?.ToRgbHex(), item.StrokeWidth, item.StrokeDashStyle,
                    item.PointColors?.Select(color => color?.ToRgbHex()).ToList(), item.ShowMarkers,
                    item.ConnectLine, item.MarkerSize, item.MarkerShape, item.MarkerOutlineColor?.ToRgbHex(),
                    item.MarkerOutlineWidth);
                series.Add(mapped);
            }
            return new ExcelChartData(data.Categories, series);
        }

        private static ExcelChartType MapChartKind(OfficeChartKind kind) {
            switch (kind) {
                case OfficeChartKind.ColumnClustered: return ExcelChartType.ColumnClustered;
                case OfficeChartKind.ColumnStacked: return ExcelChartType.ColumnStacked;
                case OfficeChartKind.ColumnStacked100: return ExcelChartType.ColumnStacked100;
                case OfficeChartKind.BarClustered: return ExcelChartType.BarClustered;
                case OfficeChartKind.BarStacked: return ExcelChartType.BarStacked;
                case OfficeChartKind.BarStacked100: return ExcelChartType.BarStacked100;
                case OfficeChartKind.Line: return ExcelChartType.Line;
                case OfficeChartKind.LineStacked: return ExcelChartType.LineStacked;
                case OfficeChartKind.LineStacked100: return ExcelChartType.LineStacked100;
                case OfficeChartKind.Area: return ExcelChartType.Area;
                case OfficeChartKind.AreaStacked: return ExcelChartType.AreaStacked;
                case OfficeChartKind.AreaStacked100: return ExcelChartType.AreaStacked100;
                case OfficeChartKind.Scatter: return ExcelChartType.Scatter;
                case OfficeChartKind.Radar: return ExcelChartType.Radar;
                case OfficeChartKind.Pie: return ExcelChartType.Pie;
                case OfficeChartKind.Doughnut: return ExcelChartType.Doughnut;
                default: return ExcelChartType.ColumnClustered;
            }
        }
    }
}
