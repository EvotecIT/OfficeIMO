namespace OfficeIMO.Internal;

/// <summary>Writes stable chart identifiers when an enum value has a convenience alias.</summary>
internal static class OfficeChartTypeName {
    internal static string Format(OfficeChartType type) => type switch {
        OfficeChartType.ColumnClustered => nameof(OfficeChartType.ColumnClustered),
        OfficeChartType.BarClustered => nameof(OfficeChartType.BarClustered),
        _ => type.ToString()
    };
}
