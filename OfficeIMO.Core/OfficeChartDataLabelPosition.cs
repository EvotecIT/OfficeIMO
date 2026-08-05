namespace OfficeIMO.Drawing;

/// <summary>
/// Preferred placement for chart data labels in reusable Office chart drawing.
/// </summary>
public enum OfficeChartDataLabelPosition {
    /// <summary>Let the renderer choose a readable default for the chart family.</summary>
    Auto,

    /// <summary>Place the label centered on the data point or mark.</summary>
    Center,

    /// <summary>Place the label inside the mark near the value axis baseline.</summary>
    InsideBase,

    /// <summary>Place the label inside the mark near the value end.</summary>
    InsideEnd,

    /// <summary>Place the label outside the mark near the value end.</summary>
    OutsideEnd,

    /// <summary>Place the label to the left of the point or mark.</summary>
    Left,

    /// <summary>Place the label to the right of the point or mark.</summary>
    Right,

    /// <summary>Place the label above the point or mark.</summary>
    Top,

    /// <summary>Place the label below the point or mark.</summary>
    Bottom
}
