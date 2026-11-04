namespace OfficeIMO.Excel;

public partial class ExcelSheet {
    private static void ValidateChartTypes(OfficeChartType type, ExcelChartData? data) {
        ValidateGeneralChartType(type);
        if (data == null) return;
        foreach (var series in data.Series) ValidateGeneralChartType(series.ChartType ?? type);
    }

    private static void ValidateGeneralChartType(OfficeChartType type) {
        if (type == OfficeChartType.Bubble)
            throw new NotSupportedException("Bubble charts require explicit X/Y/size ranges. Use AddBubbleChartFromRanges.");
        bool supported = type is
            OfficeChartType.ColumnClustered or
            OfficeChartType.ColumnStacked or
            OfficeChartType.ColumnStacked100 or
            OfficeChartType.Column3DClustered or
            OfficeChartType.Column3DStacked or
            OfficeChartType.Column3DStacked100 or
            OfficeChartType.BarClustered or
            OfficeChartType.BarStacked or
            OfficeChartType.BarStacked100 or
            OfficeChartType.Bar3DClustered or
            OfficeChartType.Bar3DStacked or
            OfficeChartType.Bar3DStacked100 or
            OfficeChartType.Line or
            OfficeChartType.LineStacked or
            OfficeChartType.LineStacked100 or
            OfficeChartType.Line3D or
            OfficeChartType.Area or
            OfficeChartType.AreaStacked or
            OfficeChartType.AreaStacked100 or
            OfficeChartType.Area3D or
            OfficeChartType.Area3DStacked or
            OfficeChartType.Area3DStacked100 or
            OfficeChartType.Pie or
            OfficeChartType.Pie3D or
            OfficeChartType.PieOfPie or
            OfficeChartType.BarOfPie or
            OfficeChartType.Doughnut or
            OfficeChartType.Scatter or
            OfficeChartType.Radar or
            OfficeChartType.Stock or
            OfficeChartType.Surface or
            OfficeChartType.SurfaceWireframe or
            OfficeChartType.SurfaceContour or
            OfficeChartType.SurfaceContourWireframe;
        if (!supported) throw new NotSupportedException($"Chart type '{type}' is not supported by the native Excel chart writer.");
    }
}
