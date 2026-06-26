namespace OfficeIMO.Excel.LegacyXls.Model {
    /// <summary>
    /// Represents formatting and visibility metadata for a legacy XLS row.
    /// </summary>
    public sealed class LegacyXlsRowLayout {
        /// <summary>
        /// Creates legacy row layout metadata.
        /// </summary>
        /// <param name="row">One-based row index.</param>
        /// <param name="height">Row height in points.</param>
        /// <param name="hidden">Whether the row is hidden.</param>
        /// <param name="customHeight">Whether the height was manually set.</param>
        /// <param name="styleIndex">Default legacy XF style index for the row, when present.</param>
        /// <param name="outlineLevel">Excel outline level from 0 through 7.</param>
        /// <param name="collapsed">Whether the row is shown as collapsed.</param>
        public LegacyXlsRowLayout(int row, double height, bool hidden, bool customHeight, ushort? styleIndex = null, byte outlineLevel = 0, bool collapsed = false) {
            Row = row;
            Height = height;
            Hidden = hidden;
            CustomHeight = customHeight;
            StyleIndex = styleIndex;
            OutlineLevel = outlineLevel;
            Collapsed = collapsed;
        }

        /// <summary>
        /// Gets the one-based row index.
        /// </summary>
        public int Row { get; }

        /// <summary>
        /// Gets the row height in points.
        /// </summary>
        public double Height { get; }

        /// <summary>
        /// Gets whether the row is hidden.
        /// </summary>
        public bool Hidden { get; }

        /// <summary>
        /// Gets whether the height was manually set.
        /// </summary>
        public bool CustomHeight { get; }

        /// <summary>
        /// Gets the default legacy XF style index for the row, when present.
        /// </summary>
        public ushort? StyleIndex { get; }

        /// <summary>
        /// Gets the Excel outline level from 0 through 7.
        /// </summary>
        public byte OutlineLevel { get; }

        /// <summary>
        /// Gets whether the row is shown as collapsed.
        /// </summary>
        public bool Collapsed { get; }
    }
}
