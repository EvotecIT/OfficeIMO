namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        /// <summary>Returns the authored pixel dimensions of an A1 range using current column widths, row heights and hidden cells.</summary>
        /// <remarks>Uses the same font-aware sizing as image range anchors. Hidden-only ranges retain the one-pixel anchor minimum.</remarks>
        public (int WidthPixels, int HeightPixels) GetRangeSizePixels(string range) {
            var bounds = ParseImageRange(range);
            return CalculateRangeAnchorSizePixels(bounds.StartRow, bounds.StartColumn, bounds.EndRow, bounds.EndColumn,
                0, 0, 0, 0);
        }
    }
}
