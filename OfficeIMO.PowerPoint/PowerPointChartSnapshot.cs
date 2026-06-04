using System;

namespace OfficeIMO.PowerPoint {
    /// <summary>
    /// PowerPoint chart families available as first-party render snapshots.
    /// </summary>
    public enum PowerPointChartSnapshotKind {
        /// <summary>Clustered column chart.</summary>
        ClusteredColumn,

        /// <summary>Stacked vertical column chart.</summary>
        StackedColumn,

        /// <summary>One-hundred percent stacked vertical column chart.</summary>
        StackedColumn100,

        /// <summary>Clustered horizontal bar chart.</summary>
        ClusteredBar,

        /// <summary>Stacked horizontal bar chart.</summary>
        StackedBar,

        /// <summary>One-hundred percent stacked horizontal bar chart.</summary>
        StackedBar100,

        /// <summary>Line chart.</summary>
        Line,

        /// <summary>Scatter chart.</summary>
        Scatter,

        /// <summary>Pie chart.</summary>
        Pie,

        /// <summary>Doughnut chart.</summary>
        Doughnut
    }

    /// <summary>
    /// Lightweight chart snapshot that consumers can render without depending on PowerPoint or Open XML chart internals.
    /// </summary>
    public sealed class PowerPointChartSnapshot {
        internal PowerPointChartSnapshot(string name, string? title, PowerPointChartSnapshotKind chartKind, PowerPointChartData data, double widthPoints, double heightPoints) {
            if (data == null) {
                throw new ArgumentNullException(nameof(data));
            }

            Name = name ?? string.Empty;
            Title = title;
            ChartKind = chartKind;
            Data = data;
            WidthPoints = widthPoints;
            HeightPoints = heightPoints;
        }

        /// <summary>Chart drawing name.</summary>
        public string Name { get; }

        /// <summary>Chart title text when present.</summary>
        public string? Title { get; }

        /// <summary>Detected chart family.</summary>
        public PowerPointChartSnapshotKind ChartKind { get; }

        /// <summary>Cached chart data.</summary>
        public PowerPointChartData Data { get; }

        /// <summary>Chart width in points.</summary>
        public double WidthPoints { get; }

        /// <summary>Chart height in points.</summary>
        public double HeightPoints { get; }
    }
}
