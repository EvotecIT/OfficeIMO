namespace OfficeIMO.Excel.LegacyXls.Model {
    /// <summary>
    /// Represents active-cell and selected-range metadata decoded from a BIFF Selection record.
    /// </summary>
    public sealed class LegacyXlsSelection {
        /// <summary>
        /// Creates a decoded worksheet selection.
        /// </summary>
        public LegacyXlsSelection(byte pane, int activeRow, int activeColumn, ushort activeRangeIndex, IReadOnlyList<LegacyXlsSelectedRange> selectedRanges) {
            Pane = pane;
            ActiveRow = activeRow;
            ActiveColumn = activeColumn;
            ActiveRangeIndex = activeRangeIndex;
            SelectedRanges = selectedRanges ?? throw new ArgumentNullException(nameof(selectedRanges));
        }

        /// <summary>Gets the legacy pane identifier.</summary>
        public byte Pane { get; }

        /// <summary>Gets the one-based active row.</summary>
        public int ActiveRow { get; }

        /// <summary>Gets the one-based active column.</summary>
        public int ActiveColumn { get; }

        /// <summary>Gets the zero-based active range index from the legacy record.</summary>
        public ushort ActiveRangeIndex { get; }

        /// <summary>Gets selected cell ranges.</summary>
        public IReadOnlyList<LegacyXlsSelectedRange> SelectedRanges { get; }
    }
}
