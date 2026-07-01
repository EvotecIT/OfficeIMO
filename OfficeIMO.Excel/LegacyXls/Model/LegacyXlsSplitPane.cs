namespace OfficeIMO.Excel.LegacyXls.Model {
    /// <summary>
    /// Represents non-frozen split pane metadata parsed from a legacy XLS worksheet view.
    /// </summary>
    public sealed class LegacyXlsSplitPane {
        /// <summary>
        /// Creates legacy split pane metadata.
        /// </summary>
        /// <param name="horizontalSplit">Horizontal split position in BIFF pane units.</param>
        /// <param name="verticalSplit">Vertical split position in BIFF pane units.</param>
        /// <param name="topRow">Top visible row of the bottom pane, zero-based.</param>
        /// <param name="leftColumn">Left visible column of the right pane, zero-based.</param>
        /// <param name="activePane">BIFF active-pane code.</param>
        public LegacyXlsSplitPane(ushort horizontalSplit, ushort verticalSplit, ushort topRow, ushort leftColumn, byte activePane) {
            HorizontalSplit = horizontalSplit;
            VerticalSplit = verticalSplit;
            TopRow = topRow;
            LeftColumn = leftColumn;
            ActivePane = activePane;
        }

        /// <summary>
        /// Gets the horizontal split position in BIFF pane units.
        /// </summary>
        public ushort HorizontalSplit { get; }

        /// <summary>
        /// Gets the vertical split position in BIFF pane units.
        /// </summary>
        public ushort VerticalSplit { get; }

        /// <summary>
        /// Gets the top visible row of the bottom pane, zero-based.
        /// </summary>
        public ushort TopRow { get; }

        /// <summary>
        /// Gets the left visible column of the right pane, zero-based.
        /// </summary>
        public ushort LeftColumn { get; }

        /// <summary>
        /// Gets the BIFF active-pane code.
        /// </summary>
        public byte ActivePane { get; }
    }
}
