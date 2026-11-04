namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        /// <summary>Returns the authored dimensions of an A1 range using current column widths, row heights and hidden cells.</summary>
        /// <remarks>Uses the same font-aware sizing as image range anchors. Hidden-only ranges retain a one-EMU native anchor minimum.</remarks>
        public (OfficeLength Width, OfficeLength Height) GetRangeSize(string range) {
            var bounds = ParseImageRange(range);
            return CalculateRangeAnchorSize(bounds.StartRow, bounds.StartColumn, bounds.EndRow, bounds.EndColumn,
                default, default, default, default);
        }
    }
}
