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
            int widthPixels = 640, int heightPixels = 360, string? title = null) {
            if (data == null) throw new ArgumentNullException(nameof(data));
            ExcelChartData excelData = ToExcelChartData(data, chartKind);
            ExcelChart chart = AddChart(excelData, row, column, widthPixels, heightPixels,
                MapChartKind(chartKind), title);
            chart.ApplyAuthoredSeriesStyles(excelData.Series,
                data.Series.Select(series => series.ShowInLegend).ToList());
            chart.ApplySharedPointStyles(data.Series);
            chart.ApplySharedPointExplosions(data.Series);
            return chart;
        }

        private static ExcelChartData ToExcelChartData(OfficeChartData data, OfficeChartType defaultKind) {
            var series = new List<ExcelChartSeries>(data.Series.Count);
            foreach (OfficeSeries item in data.Series) {
                ExcelChartType? chartType = MapChartKind(item.RenderKind ?? defaultKind);
                if (item.PointExplosions != null &&
                    chartType is not ExcelChartType.Pie and not ExcelChartType.Doughnut)
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

        private static ExcelChartType MapChartKind(OfficeChartType kind) {
            switch (kind) {
                case OfficeChartType.ColumnClustered: return ExcelChartType.ColumnClustered;
                case OfficeChartType.ColumnStacked: return ExcelChartType.ColumnStacked;
                case OfficeChartType.ColumnStacked100: return ExcelChartType.ColumnStacked100;
                case OfficeChartType.BarClustered: return ExcelChartType.BarClustered;
                case OfficeChartType.BarStacked: return ExcelChartType.BarStacked;
                case OfficeChartType.BarStacked100: return ExcelChartType.BarStacked100;
                case OfficeChartType.Line: return ExcelChartType.Line;
                case OfficeChartType.LineStacked: return ExcelChartType.LineStacked;
                case OfficeChartType.LineStacked100: return ExcelChartType.LineStacked100;
                case OfficeChartType.Area: return ExcelChartType.Area;
                case OfficeChartType.AreaStacked: return ExcelChartType.AreaStacked;
                case OfficeChartType.AreaStacked100: return ExcelChartType.AreaStacked100;
                case OfficeChartType.Scatter: return ExcelChartType.Scatter;
                case OfficeChartType.Radar: return ExcelChartType.Radar;
                case OfficeChartType.Pie: return ExcelChartType.Pie;
                case OfficeChartType.Doughnut: return ExcelChartType.Doughnut;
                case OfficeChartType.Bubble:
                    throw new NotSupportedException(
                        "Bubble charts are not supported by the shared Excel chart API. Use the native Excel bubble-chart API.");
                case OfficeChartType.Column3DClustered: return ExcelChartType.Column3DClustered;
                case OfficeChartType.Column3DStacked: return ExcelChartType.Column3DStacked;
                case OfficeChartType.Column3DStacked100: return ExcelChartType.Column3DStacked100;
                case OfficeChartType.Bar3DClustered: return ExcelChartType.Bar3DClustered;
                case OfficeChartType.Bar3DStacked: return ExcelChartType.Bar3DStacked;
                case OfficeChartType.Bar3DStacked100: return ExcelChartType.Bar3DStacked100;
                case OfficeChartType.Line3D: return ExcelChartType.Line3D;
                case OfficeChartType.Area3D: return ExcelChartType.Area3D;
                case OfficeChartType.Area3DStacked: return ExcelChartType.Area3DStacked;
                case OfficeChartType.Area3DStacked100: return ExcelChartType.Area3DStacked100;
                case OfficeChartType.Pie3D: return ExcelChartType.Pie3D;
                case OfficeChartType.PieOfPie: return ExcelChartType.PieOfPie;
                case OfficeChartType.BarOfPie: return ExcelChartType.BarOfPie;
                case OfficeChartType.Stock: return ExcelChartType.Stock;
                case OfficeChartType.Surface: return ExcelChartType.Surface;
                case OfficeChartType.SurfaceWireframe: return ExcelChartType.SurfaceWireframe;
                case OfficeChartType.SurfaceContour: return ExcelChartType.SurfaceContour;
                case OfficeChartType.SurfaceContourWireframe: return ExcelChartType.SurfaceContourWireframe;
                default:
                    throw new NotSupportedException(
                        $"Shared chart kind '{kind}' is not supported by Excel.");
            }
        }
    }
}
