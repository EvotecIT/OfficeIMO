using OfficeIMO.Drawing;

namespace OfficeIMO.Markup;

/// <summary>
/// Resolves markup chart tokens to the shared chart contract used by PowerPoint surfaces.
/// </summary>
internal static class OfficeMarkupChartKindResolver {
    internal static OfficeChartType Resolve(string chartType) {
        string normalized = new string((chartType ?? string.Empty)
            .Where(char.IsLetterOrDigit).Select(char.ToLowerInvariant).ToArray());
        return normalized switch {
            "line" => OfficeChartType.Line,
            "bar" or "clusteredbar" => OfficeChartType.BarClustered,
            "stackedbar" => OfficeChartType.BarStacked,
            "stackedcolumn" => OfficeChartType.ColumnStacked,
            "pie" => OfficeChartType.Pie,
            "doughnut" or "donut" => OfficeChartType.Doughnut,
            "scatter" => OfficeChartType.Scatter,
            "area" => OfficeChartType.Area,
            _ => OfficeChartType.ColumnClustered
        };
    }
}
