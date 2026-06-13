using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;

namespace OfficeIMO.Word {
    /// <summary>
    /// Supported chart kinds for a dependency-free Word chart snapshot.
    /// </summary>
    public enum WordChartSnapshotKind {
        /// <summary>Clustered vertical column chart.</summary>
        ClusteredColumn,
        /// <summary>Stacked vertical column chart.</summary>
        StackedColumn,
        /// <summary>One hundred percent stacked vertical column chart.</summary>
        StackedColumn100,
        /// <summary>Clustered horizontal bar chart.</summary>
        ClusteredBar,
        /// <summary>Stacked horizontal bar chart.</summary>
        StackedBar,
        /// <summary>One hundred percent stacked horizontal bar chart.</summary>
        StackedBar100,
        /// <summary>Line chart.</summary>
        Line,
        /// <summary>Stacked line chart.</summary>
        StackedLine,
        /// <summary>One hundred percent stacked line chart.</summary>
        StackedLine100,
        /// <summary>Area chart.</summary>
        Area,
        /// <summary>Stacked area chart.</summary>
        StackedArea,
        /// <summary>One hundred percent stacked area chart.</summary>
        StackedArea100,
        /// <summary>Radar chart.</summary>
        Radar,
        /// <summary>Scatter chart.</summary>
        Scatter,
        /// <summary>Pie chart.</summary>
        Pie,
        /// <summary>Doughnut chart.</summary>
        Doughnut
    }

    /// <summary>
    /// Series values extracted from cached Word chart data.
    /// </summary>
    public sealed class WordChartSeries {
        internal WordChartSeries(string name, IReadOnlyList<double> values, IReadOnlyList<double>? xValues = null, OfficeIMO.Drawing.OfficeColor? color = null, IReadOnlyList<OfficeIMO.Drawing.OfficeColor?>? pointColors = null) {
            Name = name ?? string.Empty;
            Values = new ReadOnlyCollection<double>(new List<double>(values ?? Array.Empty<double>()));
            XValues = xValues == null ? null : new ReadOnlyCollection<double>(new List<double>(xValues));
            Color = color;
            PointColors = pointColors == null ? null : new ReadOnlyCollection<OfficeIMO.Drawing.OfficeColor?>(new List<OfficeIMO.Drawing.OfficeColor?>(pointColors));
        }

        /// <summary>Series display name.</summary>
        public string Name { get; }

        /// <summary>Y values or category values.</summary>
        public IReadOnlyList<double> Values { get; }

        /// <summary>Optional X values for scatter charts.</summary>
        public IReadOnlyList<double>? XValues { get; }

        /// <summary>Optional explicit series color extracted from Word chart shape properties.</summary>
        public OfficeIMO.Drawing.OfficeColor? Color { get; }

        /// <summary>Optional explicit point colors extracted from Word chart data point shape properties.</summary>
        public IReadOnlyList<OfficeIMO.Drawing.OfficeColor?>? PointColors { get; }
    }

    /// <summary>
    /// Category and series data extracted from cached Word chart data.
    /// </summary>
    public sealed class WordChartData {
        internal WordChartData(IReadOnlyList<string> categories, IReadOnlyList<WordChartSeries> series) {
            Categories = new ReadOnlyCollection<string>(new List<string>(categories ?? Array.Empty<string>()));
            Series = new ReadOnlyCollection<WordChartSeries>(new List<WordChartSeries>(series ?? Array.Empty<WordChartSeries>()));
        }

        /// <summary>Category labels.</summary>
        public IReadOnlyList<string> Categories { get; }

        /// <summary>Chart series.</summary>
        public IReadOnlyList<WordChartSeries> Series { get; }
    }

    /// <summary>
    /// Dependency-free Word chart snapshot suitable for export and visual fallback renderers.
    /// </summary>
    public sealed class WordChartSnapshot {
        internal WordChartSnapshot(string name, string? title, WordChartSnapshotKind chartKind, WordChartData data, double widthPoints, double heightPoints) {
            Name = name ?? string.Empty;
            Title = title;
            ChartKind = chartKind;
            Data = data ?? throw new ArgumentNullException(nameof(data));
            WidthPoints = widthPoints;
            HeightPoints = heightPoints;
        }

        /// <summary>Chart drawing name when available.</summary>
        public string Name { get; }

        /// <summary>Chart title when available.</summary>
        public string? Title { get; }

        /// <summary>Detected chart kind.</summary>
        public WordChartSnapshotKind ChartKind { get; }

        /// <summary>Chart category and series data.</summary>
        public WordChartData Data { get; }

        /// <summary>Chart frame width in points.</summary>
        public double WidthPoints { get; }

        /// <summary>Chart frame height in points.</summary>
        public double HeightPoints { get; }
    }
}
