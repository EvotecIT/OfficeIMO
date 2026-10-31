namespace OfficeIMO;

/// <summary>
/// Chart families shared by document authoring formats.
/// </summary>
public enum OfficeChartType {
    /// <summary>Clustered vertical column chart.</summary>
    ColumnClustered,

    /// <summary>Ordinary clustered column chart.</summary>
    Column = ColumnClustered,

    /// <summary>Stacked vertical column chart.</summary>
    ColumnStacked,

    /// <summary>One-hundred percent stacked vertical column chart.</summary>
    ColumnStacked100,

    /// <summary>Clustered horizontal bar chart.</summary>
    BarClustered,

    /// <summary>Ordinary clustered bar chart.</summary>
    Bar = BarClustered,

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
    Doughnut,

    /// <summary>Bubble chart with numeric X/Y coordinates and per-point sizes.</summary>
    Bubble,

    /// <summary>3-D clustered column (vertical bars).</summary>
    Column3DClustered,

    /// <summary>3-D stacked column (vertical bars).</summary>
    Column3DStacked,

    /// <summary>3-D 100% stacked column (vertical bars).</summary>
    Column3DStacked100,

    /// <summary>3-D clustered bar (horizontal bars).</summary>
    Bar3DClustered,

    /// <summary>3-D stacked bar (horizontal bars).</summary>
    Bar3DStacked,

    /// <summary>3-D 100% stacked bar (horizontal bars).</summary>
    Bar3DStacked100,

    /// <summary>3-D line chart.</summary>
    Line3D,

    /// <summary>3-D area chart.</summary>
    Area3D,

    /// <summary>3-D stacked area chart.</summary>
    Area3DStacked,

    /// <summary>3-D 100% stacked area chart.</summary>
    Area3DStacked100,

    /// <summary>3-D pie chart.</summary>
    Pie3D,

    /// <summary>Pie-of-pie chart.</summary>
    PieOfPie,

    /// <summary>Bar-of-pie chart.</summary>
    BarOfPie,

    /// <summary>Stock chart using high-low-close or open-high-low-close series.</summary>
    Stock,

    /// <summary>3-D surface chart.</summary>
    Surface,

    /// <summary>Wireframe 3-D surface chart.</summary>
    SurfaceWireframe,

    /// <summary>Contour surface chart.</summary>
    SurfaceContour,

    /// <summary>Native funnel chart.</summary>
    Funnel,

    /// <summary>Native waterfall chart.</summary>
    Waterfall,

    /// <summary>Native box-and-whisker chart.</summary>
    BoxWhisker,

    /// <summary>Native treemap chart.</summary>
    Treemap,

    /// <summary>Wireframe contour surface chart.</summary>
    SurfaceContourWireframe,

    /// <summary>Native sunburst chart.</summary>
    Sunburst
}
