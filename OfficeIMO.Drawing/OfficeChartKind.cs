namespace OfficeIMO.Drawing;

/// <summary>
/// Chart families understood by shared OfficeIMO drawing renderers.
/// </summary>
public enum OfficeChartKind {
    /// <summary>Clustered vertical column chart.</summary>
    ColumnClustered,

    /// <summary>Stacked vertical column chart.</summary>
    ColumnStacked,

    /// <summary>One-hundred percent stacked vertical column chart.</summary>
    ColumnStacked100,

    /// <summary>Clustered horizontal bar chart.</summary>
    BarClustered,

    /// <summary>Stacked horizontal bar chart.</summary>
    BarStacked,

    /// <summary>One-hundred percent stacked horizontal bar chart.</summary>
    BarStacked100,

    /// <summary>Line chart.</summary>
    Line,

    /// <summary>Stacked line chart.</summary>
    LineStacked,

    /// <summary>One-hundred percent stacked line chart.</summary>
    LineStacked100,

    /// <summary>Area chart.</summary>
    Area,

    /// <summary>Stacked area chart.</summary>
    AreaStacked,

    /// <summary>One-hundred percent stacked area chart.</summary>
    AreaStacked100,

    /// <summary>Scatter chart with markers and connecting lines.</summary>
    Scatter,

    /// <summary>Radar chart.</summary>
    Radar,

    /// <summary>Pie chart.</summary>
    Pie,

    /// <summary>Doughnut chart.</summary>
    Doughnut
}
