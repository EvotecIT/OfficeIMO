using OfficeIMO.Drawing;

namespace OfficeIMO.Markup;

/// <summary>
/// Resolves markup chart tokens to the shared chart contract used by PowerPoint surfaces.
/// </summary>
internal static class OfficeMarkupChartKindResolver {
    internal static OfficeChartKind Resolve(string chartType) {
        string normalized = new string((chartType ?? string.Empty)
            .Where(char.IsLetterOrDigit).Select(char.ToLowerInvariant).ToArray());
        return normalized switch {
            "line" => OfficeChartKind.Line,
            "bar" or "clusteredbar" => OfficeChartKind.BarClustered,
            "stackedbar" => OfficeChartKind.BarStacked,
            "stackedcolumn" => OfficeChartKind.ColumnStacked,
            "pie" => OfficeChartKind.Pie,
            "doughnut" or "donut" => OfficeChartKind.Doughnut,
            "scatter" => OfficeChartKind.Scatter,
            "area" => OfficeChartKind.Area,
            _ => OfficeChartKind.ColumnClustered
        };
    }
}
