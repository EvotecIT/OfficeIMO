namespace OfficeIMO.Drawing;

public sealed partial class OfficeChartLayout {
    /// <summary>Whether outside radial labels connect to their slices.</summary>
    public bool ShowDataLabelLeaderLines { get; private set; } = true;

    /// <summary>Returns a layout copy with the requested outside radial leader-line setting.</summary>
    public OfficeChartLayout WithDataLabelLeaderLines(bool show) {
        var copy = (OfficeChartLayout)MemberwiseClone();
        copy.ShowDataLabelLeaderLines = show;
        return copy;
    }
}
