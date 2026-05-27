using Color = OfficeIMO.Drawing.OfficeColor;

namespace OfficeIMO.Visio.Diagrams {
    /// <summary>
    /// Visual defaults used by <see cref="VisioFlowchartBuilder"/>.
    /// </summary>
    public sealed class VisioFlowchartTheme {
        /// <summary>Default process fill.</summary>
        public Color ProcessFill { get; set; } = Color.FromRgb(45, 145, 225);

        /// <summary>Default process stroke.</summary>
        public Color ProcessStroke { get; set; } = Color.FromRgb(26, 111, 188);

        /// <summary>Default decision fill.</summary>
        public Color DecisionFill { get; set; } = Color.FromRgb(120, 199, 89);

        /// <summary>Default decision stroke.</summary>
        public Color DecisionStroke { get; set; } = Color.FromRgb(93, 166, 68);

        /// <summary>Default terminator fill.</summary>
        public Color TerminatorFill { get; set; } = Color.FromRgb(122, 202, 94);

        /// <summary>Default terminator stroke.</summary>
        public Color TerminatorStroke { get; set; } = Color.FromRgb(93, 166, 68);

        /// <summary>Default continuation marker fill.</summary>
        public Color MarkerFill { get; set; } = Color.FromRgb(20, 119, 211);

        /// <summary>Default continuation marker stroke.</summary>
        public Color MarkerStroke { get; set; } = Color.FromRgb(16, 96, 170);

        /// <summary>Default connector color.</summary>
        public Color ConnectorColor { get; set; } = Color.FromRgb(45, 145, 225);

        /// <summary>Default line weight for shapes and connectors, in Visio internal units.</summary>
        public double LineWeight { get; set; } = 0.018;

        /// <summary>Default process width in page units.</summary>
        public double ProcessWidth { get; set; } = 2.35;

        /// <summary>Default process height in page units.</summary>
        public double ProcessHeight { get; set; } = 1.05;

        /// <summary>Default process text style.</summary>
        public VisioTextStyle? ProcessTextStyle { get; set; }

        /// <summary>Default decision width in page units.</summary>
        public double DecisionWidth { get; set; } = 2.35;

        /// <summary>Default decision height in page units.</summary>
        public double DecisionHeight { get; set; } = 1.45;

        /// <summary>Default decision text style.</summary>
        public VisioTextStyle? DecisionTextStyle { get; set; }

        /// <summary>Default terminator width in page units.</summary>
        public double TerminatorWidth { get; set; } = 2.35;

        /// <summary>Default terminator height in page units.</summary>
        public double TerminatorHeight { get; set; } = 0.9;

        /// <summary>Default terminator text style.</summary>
        public VisioTextStyle? TerminatorTextStyle { get; set; }

        /// <summary>Default continuation marker diameter in page units.</summary>
        public double MarkerDiameter { get; set; } = 0.55;

        /// <summary>Default continuation marker text style.</summary>
        public VisioTextStyle? MarkerTextStyle { get; set; }

        /// <summary>Default connector label text style.</summary>
        public VisioTextStyle? ConnectorTextStyle { get; set; }

        /// <summary>Default title text style.</summary>
        public VisioTextStyle? TitleTextStyle { get; set; } = new VisioTextStyle {
            FontFamily = "Aptos Display",
            Color = Color.FromRgb(0, 0, 0),
            Size = 20,
            Bold = true,
            HorizontalAlignment = VisioTextHorizontalAlignment.Center,
            VerticalAlignment = VisioTextVerticalAlignment.Middle
        };

        /// <summary>Creates a detached copy of the theme.</summary>
        public VisioFlowchartTheme Clone() => new VisioFlowchartTheme {
            ProcessFill = ProcessFill,
            ProcessStroke = ProcessStroke,
            DecisionFill = DecisionFill,
            DecisionStroke = DecisionStroke,
            TerminatorFill = TerminatorFill,
            TerminatorStroke = TerminatorStroke,
            MarkerFill = MarkerFill,
            MarkerStroke = MarkerStroke,
            ConnectorColor = ConnectorColor,
            LineWeight = LineWeight,
            ProcessWidth = ProcessWidth,
            ProcessHeight = ProcessHeight,
            DecisionWidth = DecisionWidth,
            DecisionHeight = DecisionHeight,
            TerminatorWidth = TerminatorWidth,
            TerminatorHeight = TerminatorHeight,
            MarkerDiameter = MarkerDiameter,
            ProcessTextStyle = ProcessTextStyle?.Clone(),
            DecisionTextStyle = DecisionTextStyle?.Clone(),
            TerminatorTextStyle = TerminatorTextStyle?.Clone(),
            MarkerTextStyle = MarkerTextStyle?.Clone(),
            ConnectorTextStyle = ConnectorTextStyle?.Clone(),
            TitleTextStyle = TitleTextStyle?.Clone()
        };

        /// <summary>Default blue/green OfficeIMO flowchart theme.</summary>
        public static VisioFlowchartTheme ModernBlueGreen() => new VisioFlowchartTheme();
    }
}
