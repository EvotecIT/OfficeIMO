namespace OfficeIMO.Drawing;

public sealed partial class OfficeChartLayout {
    /// <summary>Independent secondary value-axis scale and format, or automatic settings when absent.</summary>
    public OfficeChartValueAxisLayout? SecondaryValueAxis { get; private set; }

    /// <summary>Returns a layout copy with an independent secondary value-axis scale and format.</summary>
    public OfficeChartLayout WithSecondaryValueAxis(OfficeChartValueAxisLayout? axis) {
        var copy = (OfficeChartLayout)MemberwiseClone();
        copy.SecondaryValueAxis = axis;
        return copy;
    }
}
