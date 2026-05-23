namespace OfficeIMO.Excel {
    /// <summary>
    /// Supported Excel chart types for OfficeIMO.Excel chart helpers.
    /// </summary>
    public enum ExcelChartType {
        /// <summary>Clustered column (vertical bars).</summary>
        ColumnClustered,
        /// <summary>Stacked column (vertical bars).</summary>
        ColumnStacked,
        /// <summary>100% stacked column (vertical bars).</summary>
        ColumnStacked100,
        /// <summary>3-D clustered column (vertical bars).</summary>
        Column3DClustered,
        /// <summary>3-D stacked column (vertical bars).</summary>
        Column3DStacked,
        /// <summary>3-D 100% stacked column (vertical bars).</summary>
        Column3DStacked100,
        /// <summary>Clustered bar (horizontal bars).</summary>
        BarClustered,
        /// <summary>Stacked bar (horizontal bars).</summary>
        BarStacked,
        /// <summary>100% stacked bar (horizontal bars).</summary>
        BarStacked100,
        /// <summary>3-D clustered bar (horizontal bars).</summary>
        Bar3DClustered,
        /// <summary>3-D stacked bar (horizontal bars).</summary>
        Bar3DStacked,
        /// <summary>3-D 100% stacked bar (horizontal bars).</summary>
        Bar3DStacked100,
        /// <summary>Line chart.</summary>
        Line,
        /// <summary>Stacked line chart.</summary>
        LineStacked,
        /// <summary>100% stacked line chart.</summary>
        LineStacked100,
        /// <summary>3-D line chart.</summary>
        Line3D,
        /// <summary>Area chart.</summary>
        Area,
        /// <summary>Stacked area chart.</summary>
        AreaStacked,
        /// <summary>100% stacked area chart.</summary>
        AreaStacked100,
        /// <summary>3-D area chart.</summary>
        Area3D,
        /// <summary>3-D stacked area chart.</summary>
        Area3DStacked,
        /// <summary>3-D 100% stacked area chart.</summary>
        Area3DStacked100,
        /// <summary>Pie chart.</summary>
        Pie,
        /// <summary>3-D pie chart.</summary>
        Pie3D,
        /// <summary>Pie-of-pie chart.</summary>
        PieOfPie,
        /// <summary>Bar-of-pie chart.</summary>
        BarOfPie,
        /// <summary>Doughnut chart.</summary>
        Doughnut,
        /// <summary>Scatter (XY) chart.</summary>
        Scatter,
        /// <summary>Bubble chart.</summary>
        Bubble,
        /// <summary>Radar chart.</summary>
        Radar,
        /// <summary>Stock chart using high-low-close or open-high-low-close series.</summary>
        Stock,
        /// <summary>3-D surface chart.</summary>
        Surface,
        /// <summary>Wireframe 3-D surface chart.</summary>
        SurfaceWireframe,
        /// <summary>Contour surface chart.</summary>
        SurfaceContour,
        /// <summary>Wireframe contour surface chart.</summary>
        SurfaceContourWireframe
    }
}
