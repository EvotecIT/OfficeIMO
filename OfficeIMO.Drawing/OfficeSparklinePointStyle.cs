namespace OfficeIMO.Drawing;

/// <summary>
/// Per-point sparkline rendering style resolved by a source adapter.
/// </summary>
public readonly struct OfficeSparklinePointStyle {
    /// <summary>
    /// Creates a per-point sparkline style.
    /// </summary>
    /// <param name="color">Rendered point color.</param>
    /// <param name="showMarker">Whether line sparklines should draw a marker at this point.</param>
    public OfficeSparklinePointStyle(OfficeColor color, bool showMarker = false) {
        Color = color;
        ShowMarker = showMarker;
    }

    /// <summary>Rendered point color.</summary>
    public OfficeColor Color { get; }

    /// <summary>Whether line sparklines should draw a marker at this point.</summary>
    public bool ShowMarker { get; }
}
