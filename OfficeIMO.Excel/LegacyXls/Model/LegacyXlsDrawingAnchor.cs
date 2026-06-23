namespace OfficeIMO.Excel.LegacyXls.Model {
    /// <summary>
    /// Describes shallow metadata decoded from an OfficeArtFClientAnchor record.
    /// </summary>
    public sealed class LegacyXlsDrawingAnchor {
        /// <summary>
        /// Creates preserve-only metadata for an OfficeArt client anchor.
        /// </summary>
        public LegacyXlsDrawingAnchor(
            ushort flags,
            ushort startColumn,
            ushort startDx,
            ushort startRow,
            ushort startDy,
            ushort endColumn,
            ushort endDx,
            ushort endRow,
            ushort endDy) {
            Flags = flags;
            StartColumn = startColumn;
            StartDx = startDx;
            StartRow = startRow;
            StartDy = startDy;
            EndColumn = endColumn;
            EndDx = endDx;
            EndRow = endRow;
            EndDy = endDy;
        }

        /// <summary>Gets the raw anchor flags.</summary>
        public ushort Flags { get; }

        /// <summary>Gets the zero-based starting column.</summary>
        public ushort StartColumn { get; }

        /// <summary>Gets the starting column offset.</summary>
        public ushort StartDx { get; }

        /// <summary>Gets the zero-based starting row.</summary>
        public ushort StartRow { get; }

        /// <summary>Gets the starting row offset.</summary>
        public ushort StartDy { get; }

        /// <summary>Gets the zero-based ending column.</summary>
        public ushort EndColumn { get; }

        /// <summary>Gets the ending column offset.</summary>
        public ushort EndDx { get; }

        /// <summary>Gets the zero-based ending row.</summary>
        public ushort EndRow { get; }

        /// <summary>Gets the ending row offset.</summary>
        public ushort EndDy { get; }
    }
}
