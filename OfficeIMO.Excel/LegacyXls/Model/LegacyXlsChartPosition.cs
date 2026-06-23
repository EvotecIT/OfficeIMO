namespace OfficeIMO.Excel.LegacyXls.Model {
    /// <summary>
    /// Describes shallow metadata decoded from a chart Pos record.
    /// </summary>
    public sealed class LegacyXlsChartPosition {
        internal LegacyXlsChartPosition(
            ushort topLeftMode,
            string topLeftModeName,
            ushort bottomRightMode,
            string bottomRightModeName,
            short x1,
            short y1,
            short x2,
            short y2) {
            TopLeftMode = topLeftMode;
            TopLeftModeName = topLeftModeName ?? throw new ArgumentNullException(nameof(topLeftModeName));
            BottomRightMode = bottomRightMode;
            BottomRightModeName = bottomRightModeName ?? throw new ArgumentNullException(nameof(bottomRightModeName));
            X1 = x1;
            Y1 = y1;
            X2 = x2;
            Y2 = y2;
        }

        /// <summary>Gets the raw upper-left position mode.</summary>
        public ushort TopLeftMode { get; }

        /// <summary>Gets the decoded upper-left position mode name.</summary>
        public string TopLeftModeName { get; }

        /// <summary>Gets the raw lower-right position mode.</summary>
        public ushort BottomRightMode { get; }

        /// <summary>Gets the decoded lower-right position mode name.</summary>
        public string BottomRightModeName { get; }

        /// <summary>Gets the first X coordinate or offset.</summary>
        public short X1 { get; }

        /// <summary>Gets the first Y coordinate or offset.</summary>
        public short Y1 { get; }

        /// <summary>Gets the second X coordinate, width, or ignored value.</summary>
        public short X2 { get; }

        /// <summary>Gets the second Y coordinate, height, or ignored value.</summary>
        public short Y2 { get; }
    }
}
