namespace OfficeIMO.PowerPoint;

/// <summary>Controls the title, size, placement and accessibility of an editable PowerPoint chart.</summary>
public sealed class PowerPointChartOptions : OfficeChartOptions {
    /// <summary>Gets or sets the horizontal position. Null places the chart at the slide origin.</summary>
    public OfficeLength? Left { get; set; }

    /// <summary>Gets or sets the vertical position. Null places the chart at the slide origin.</summary>
    public OfficeLength? Top { get; set; }

    /// <summary>Gets or sets accessible naming and alternative-text options.</summary>
    public PowerPointChartAccessibilityOptions? Accessibility { get; set; }
}
