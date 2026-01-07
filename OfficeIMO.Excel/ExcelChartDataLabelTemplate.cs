using DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.Excel {
    /// <summary>
    /// Defines reusable data label settings for chart series.
    /// </summary>
    public sealed class ExcelChartDataLabelTemplate {
        /// <summary>Toggle value labels.</summary>
        public bool? ShowValue { get; set; }
        /// <summary>Toggle category name labels.</summary>
        public bool? ShowCategoryName { get; set; }
        /// <summary>Toggle series name labels.</summary>
        public bool? ShowSeriesName { get; set; }
        /// <summary>Toggle legend key labels.</summary>
        public bool? ShowLegendKey { get; set; }
        /// <summary>Toggle percent labels.</summary>
        public bool? ShowPercent { get; set; }
        /// <summary>Toggle bubble size labels.</summary>
        public bool? ShowBubbleSize { get; set; }
        /// <summary>Sets label position.</summary>
        public DataLabelPositionValues? Position { get; set; }
        /// <summary>Number format code for labels.</summary>
        public string? NumberFormat { get; set; }
        /// <summary>Separator between label parts.</summary>
        public string? Separator { get; set; }
        /// <summary>Use source-linked formatting for number format.</summary>
        public bool SourceLinked { get; set; }

        /// <summary>Toggle leader lines for labels.</summary>
        public bool? ShowLeaderLines { get; set; }
        /// <summary>Leader line color (hex).</summary>
        public string? LeaderLineColor { get; set; }
        /// <summary>Leader line width in points.</summary>
        public double? LeaderLineWidthPoints { get; set; }

        /// <summary>Font size in points.</summary>
        public double? FontSizePoints { get; set; }
        /// <summary>Toggle bold text.</summary>
        public bool? Bold { get; set; }
        /// <summary>Toggle italic text.</summary>
        public bool? Italic { get; set; }
        /// <summary>Font name.</summary>
        public string? FontName { get; set; }
        /// <summary>Text color (hex).</summary>
        public string? TextColor { get; set; }

        /// <summary>Fill color (hex).</summary>
        public string? FillColor { get; set; }
        /// <summary>Line color (hex).</summary>
        public string? LineColor { get; set; }
        /// <summary>Line width in points.</summary>
        public double? LineWidthPoints { get; set; }
        /// <summary>Disable fill.</summary>
        public bool NoFill { get; set; }
        /// <summary>Disable line.</summary>
        public bool NoLine { get; set; }
    }
}
