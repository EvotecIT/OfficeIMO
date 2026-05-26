using Color = OfficeIMO.Drawing.OfficeColor;

namespace OfficeIMO.Visio.Diagrams {
    /// <summary>
    /// Visual defaults used by <see cref="VisioBlockDiagramBuilder"/>.
    /// </summary>
    public sealed class VisioBlockDiagramTheme {
        /// <summary>Default block fill color.</summary>
        public Color BlockFill { get; set; } = Color.FromRgb(0, 146, 203);

        /// <summary>Default block stroke color.</summary>
        public Color BlockStroke { get; set; } = Color.FromRgb(0, 106, 160);

        /// <summary>Default emphasis block fill color.</summary>
        public Color EmphasisFill { get; set; } = Color.FromRgb(88, 88, 88);

        /// <summary>Default emphasis block stroke color.</summary>
        public Color EmphasisStroke { get; set; } = Color.FromRgb(64, 64, 64);

        /// <summary>Default region fill color.</summary>
        public Color RegionFill { get; set; } = Color.FromRgb(218, 242, 252);

        /// <summary>Default region stroke color.</summary>
        public Color RegionStroke { get; set; } = Color.FromRgb(146, 202, 224);

        /// <summary>Solid data-flow connector color.</summary>
        public Color DataFlowColor { get; set; } = Color.FromRgb(0, 146, 203);

        /// <summary>Dashed control-flow connector color.</summary>
        public Color ControlFlowColor { get; set; } = Color.FromRgb(0, 146, 203);

        /// <summary>Default block width in page units.</summary>
        public double BlockWidth { get; set; } = 2.35;

        /// <summary>Default block height in page units.</summary>
        public double BlockHeight { get; set; } = 1.0;

        /// <summary>Default block text style.</summary>
        public VisioTextStyle? BlockTextStyle { get; set; }

        /// <summary>Default emphasis block text style.</summary>
        public VisioTextStyle? EmphasisTextStyle { get; set; }

        /// <summary>Default region text style.</summary>
        public VisioTextStyle? RegionTextStyle { get; set; }

        /// <summary>Default connector label text style.</summary>
        public VisioTextStyle? ConnectorTextStyle { get; set; }

        /// <summary>Default column gap in page units.</summary>
        public double ColumnGap { get; set; } = 1.1;

        /// <summary>Default row gap in page units.</summary>
        public double RowGap { get; set; } = 0.8;

        /// <summary>Default shape and connector line weight.</summary>
        public double LineWeight { get; set; } = 0.018;

        /// <summary>Creates a detached copy of the theme.</summary>
        public VisioBlockDiagramTheme Clone() => new VisioBlockDiagramTheme {
            BlockFill = BlockFill,
            BlockStroke = BlockStroke,
            EmphasisFill = EmphasisFill,
            EmphasisStroke = EmphasisStroke,
            RegionFill = RegionFill,
            RegionStroke = RegionStroke,
            DataFlowColor = DataFlowColor,
            ControlFlowColor = ControlFlowColor,
            BlockWidth = BlockWidth,
            BlockHeight = BlockHeight,
            ColumnGap = ColumnGap,
            RowGap = RowGap,
            LineWeight = LineWeight,
            BlockTextStyle = BlockTextStyle?.Clone(),
            EmphasisTextStyle = EmphasisTextStyle?.Clone(),
            RegionTextStyle = RegionTextStyle?.Clone(),
            ConnectorTextStyle = ConnectorTextStyle?.Clone()
        };

        /// <summary>Default blue/gray OfficeIMO block diagram theme.</summary>
        public static VisioBlockDiagramTheme TechnicalBlue() => new VisioBlockDiagramTheme();
    }
}
