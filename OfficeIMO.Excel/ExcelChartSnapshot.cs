using System;

namespace OfficeIMO.Excel {
    /// <summary>
    /// Lightweight chart snapshot that consumers can render without depending on Excel or Open XML chart internals.
    /// </summary>
    public sealed class ExcelChartSnapshot {
        internal ExcelChartSnapshot(
            string name,
            string? title,
            ExcelChartType chartType,
            ExcelChartData data,
            int rowIndex,
            int columnIndex,
            int widthPixels,
            int heightPixels) {
            Name = name ?? string.Empty;
            Title = title;
            ChartType = chartType;
            Data = data ?? throw new ArgumentNullException(nameof(data));
            RowIndex = rowIndex;
            ColumnIndex = columnIndex;
            WidthPixels = widthPixels;
            HeightPixels = heightPixels;
        }

        /// <summary>Chart drawing name.</summary>
        public string Name { get; }

        /// <summary>Chart title text when present.</summary>
        public string? Title { get; }

        /// <summary>Detected chart type.</summary>
        public ExcelChartType ChartType { get; }

        /// <summary>Cached or worksheet-backed chart data.</summary>
        public ExcelChartData Data { get; }

        /// <summary>One-based worksheet row where the chart is anchored when known.</summary>
        public int RowIndex { get; }

        /// <summary>One-based worksheet column where the chart is anchored when known.</summary>
        public int ColumnIndex { get; }

        /// <summary>Chart width in pixels when known.</summary>
        public int WidthPixels { get; }

        /// <summary>Chart height in pixels when known.</summary>
        public int HeightPixels { get; }
    }
}
