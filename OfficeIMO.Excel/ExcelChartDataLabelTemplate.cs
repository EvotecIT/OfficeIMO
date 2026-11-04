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
        public OfficeChartDataLabelPosition? Position { get; set; }
        /// <summary>Number format code for labels.</summary>
        public string? NumberFormat { get; set; }
        /// <summary>Separator between label parts.</summary>
        public string? Separator { get; set; }
        /// <summary>Use source-linked formatting for number format.</summary>
        public bool SourceLinked { get; set; }

        /// <summary>Toggle leader lines for labels.</summary>
        public bool? ShowLeaderLines { get; set; }
        /// <summary>Leader line color.</summary>
        public OfficeColor? LeaderLineColor { get; set; }
        /// <summary>Leader line physical width.</summary>
        public OfficeLength? LeaderLineWidth { get; set; }

        /// <summary>Physical font size.</summary>
        public OfficeLength? FontSize { get; set; }
        /// <summary>Toggle bold text.</summary>
        public bool? Bold { get; set; }
        /// <summary>Toggle italic text.</summary>
        public bool? Italic { get; set; }
        /// <summary>Font name.</summary>
        public string? FontName { get; set; }
        /// <summary>Text color.</summary>
        public OfficeColor? TextColor { get; set; }

        /// <summary>Fill color.</summary>
        public OfficeColor? FillColor { get; set; }
        /// <summary>Line color.</summary>
        public OfficeColor? LineColor { get; set; }
        /// <summary>Line physical width.</summary>
        public OfficeLength? LineWidth { get; set; }
        /// <summary>Disable fill.</summary>
        public bool NoFill { get; set; }
        /// <summary>Disable line.</summary>
        public bool NoLine { get; set; }
    }
}
