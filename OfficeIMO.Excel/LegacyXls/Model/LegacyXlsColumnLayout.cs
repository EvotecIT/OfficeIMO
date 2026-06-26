namespace OfficeIMO.Excel.LegacyXls.Model {
    /// <summary>
    /// Represents formatting and visibility metadata for a legacy XLS column range.
    /// </summary>
    public sealed class LegacyXlsColumnLayout {
        /// <summary>
        /// Creates legacy column layout metadata.
        /// </summary>
        /// <param name="startColumn">One-based first column.</param>
        /// <param name="endColumn">One-based last column.</param>
        /// <param name="width">Column width in Excel character-width units.</param>
        /// <param name="hidden">Whether the column range is hidden.</param>
        /// <param name="styleIndex">Default legacy XF style index for the column range.</param>
        /// <param name="outlineLevel">Excel outline level from 0 through 7.</param>
        /// <param name="collapsed">Whether the column range is shown as collapsed.</param>
        public LegacyXlsColumnLayout(int startColumn, int endColumn, double width, bool hidden, ushort styleIndex, byte outlineLevel = 0, bool collapsed = false) {
            StartColumn = startColumn;
            EndColumn = endColumn;
            Width = width;
            Hidden = hidden;
            StyleIndex = styleIndex;
            OutlineLevel = outlineLevel;
            Collapsed = collapsed;
        }

        /// <summary>
        /// Gets the one-based first column.
        /// </summary>
        public int StartColumn { get; }

        /// <summary>
        /// Gets the one-based last column.
        /// </summary>
        public int EndColumn { get; }

        /// <summary>
        /// Gets the column width in Excel character-width units.
        /// </summary>
        public double Width { get; }

        /// <summary>
        /// Gets whether the column range is hidden.
        /// </summary>
        public bool Hidden { get; }

        /// <summary>
        /// Gets the default legacy XF style index for the column range.
        /// </summary>
        public ushort StyleIndex { get; }

        /// <summary>
        /// Gets the Excel outline level from 0 through 7.
        /// </summary>
        public byte OutlineLevel { get; }

        /// <summary>
        /// Gets whether the column range is shown as collapsed.
        /// </summary>
        public bool Collapsed { get; }
    }
}
