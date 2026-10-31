using System;
using System.Collections.Generic;
using OfficeIMO.Drawing;

namespace OfficeIMO.Excel {
    /// <summary>
    /// Lightweight chart snapshot that consumers can render without depending on Excel or Open XML chart internals.
    /// </summary>
    public sealed class ExcelChartInfo {
        internal ExcelChartInfo(
            string name,
            string? title,
            OfficeChartType chartType,
            ExcelChartData data,
            int rowIndex,
            int columnIndex,
            OfficeLength offsetX,
            OfficeLength offsetY,
            OfficeLength width,
            OfficeLength height,
            OfficeChartStyle? style = null,
            OfficeChartLayout? layout = null,
            IReadOnlyList<OfficeImageExportDiagnostic>? diagnostics = null,
            OfficeChartRadialLayout? radialLayout = null,
            double bubbleScalePercent = 100D,
            OfficeChartBubbleSizeMode bubbleSizeMode = OfficeChartBubbleSizeMode.Area,
            bool showNegativeBubbles = false) {
            Name = name ?? string.Empty;
            Title = title;
            ChartType = chartType;
            Data = data ?? throw new ArgumentNullException(nameof(data));
            RowIndex = rowIndex;
            ColumnIndex = columnIndex;
            OffsetX = offsetX;
            OffsetY = offsetY;
            Width = width;
            Height = height;
            Style = style;
            Layout = layout;
            Diagnostics = diagnostics ?? Array.Empty<OfficeImageExportDiagnostic>();
            RadialLayout = radialLayout ?? OfficeChartRadialLayout.Default;
            BubbleScalePercent = bubbleScalePercent;
            BubbleSizeMode = bubbleSizeMode;
            ShowNegativeBubbles = showNegativeBubbles;
        }

        /// <summary>Chart drawing name.</summary>
        public string Name { get; }

        /// <summary>Chart title text when present.</summary>
        public string? Title { get; }

        /// <summary>Detected chart type.</summary>
        public OfficeChartType ChartType { get; }

        /// <summary>Cached or worksheet-backed chart data.</summary>
        public ExcelChartData Data { get; }

        /// <summary>One-based worksheet row where the chart is anchored when known.</summary>
        public int RowIndex { get; }

        /// <summary>One-based worksheet column where the chart is anchored when known.</summary>
        public int ColumnIndex { get; }

        /// <summary>Horizontal offset from the anchor column.</summary>
        public OfficeLength OffsetX { get; }

        /// <summary>Vertical offset from the anchor row.</summary>
        public OfficeLength OffsetY { get; }

        /// <summary>Physical chart width when known.</summary>
        public OfficeLength Width { get; }

        /// <summary>Physical chart height when known.</summary>
        public OfficeLength Height { get; }

        /// <summary>Shared rendering style metadata extracted from the chart, when available.</summary>
        public OfficeChartStyle? Style { get; }

        /// <summary>Shared rendering layout metadata extracted from the chart, when available.</summary>
        public OfficeChartLayout? Layout { get; }

        /// <summary>Diagnostics for authored chart features that the shared image renderer approximates or cannot render yet.</summary>
        public IReadOnlyList<OfficeImageExportDiagnostic> Diagnostics { get; }

        /// <summary>Authored bubble diameter scale percentage; defaults to 100.</summary>
        public double BubbleScalePercent { get; }

        /// <summary>Whether source bubble sizes represent area or width.</summary>
        public OfficeChartBubbleSizeMode BubbleSizeMode { get; }

        /// <summary>Whether the source requests negative-size bubbles. Unsupported painting is diagnosed.</summary>
        public bool ShowNegativeBubbles { get; }

        /// <summary>Native pie rotation and doughnut hole geometry.</summary>
        public OfficeChartRadialLayout RadialLayout { get; }
    }
}
