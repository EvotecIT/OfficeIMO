namespace OfficeIMO.Drawing;

/// <summary>The bounds and source value of a Cartesian data point placed by the chart renderer.</summary>
internal readonly record struct OfficeChartPointPlacement(
    double X, double Y, double Width, double Height, string Series, string? Category, double? NumericX, double Value);
