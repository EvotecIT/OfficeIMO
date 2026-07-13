namespace OfficeIMO.Visio.Diagrams {
    /// <summary>
    /// Layout dimensions used by <see cref="VisioFlowchartBuilder"/>. Visual styling is owned by
    /// <see cref="VisioStyleTheme"/>.
    /// </summary>
    public sealed class VisioFlowchartLayoutOptions {
        /// <summary>Default process width in page units.</summary>
        public double ProcessWidth { get; set; } = 2.35;

        /// <summary>Default process height in page units.</summary>
        public double ProcessHeight { get; set; } = 1.05;

        /// <summary>Default decision width in page units.</summary>
        public double DecisionWidth { get; set; } = 2.35;

        /// <summary>Default decision height in page units.</summary>
        public double DecisionHeight { get; set; } = 1.45;

        /// <summary>Default terminator width in page units.</summary>
        public double TerminatorWidth { get; set; } = 2.35;

        /// <summary>Default terminator height in page units.</summary>
        public double TerminatorHeight { get; set; } = 0.9;

        /// <summary>Default continuation marker diameter in page units.</summary>
        public double MarkerDiameter { get; set; } = 0.55;

        /// <summary>Creates a detached copy of these options.</summary>
        public VisioFlowchartLayoutOptions Clone() => new VisioFlowchartLayoutOptions {
            ProcessWidth = ProcessWidth,
            ProcessHeight = ProcessHeight,
            DecisionWidth = DecisionWidth,
            DecisionHeight = DecisionHeight,
            TerminatorWidth = TerminatorWidth,
            TerminatorHeight = TerminatorHeight,
            MarkerDiameter = MarkerDiameter
        };
    }
}
