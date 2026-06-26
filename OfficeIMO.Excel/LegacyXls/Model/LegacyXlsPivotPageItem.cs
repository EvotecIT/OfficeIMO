namespace OfficeIMO.Excel.LegacyXls.Model {
    /// <summary>
    /// Describes one SXPI page-axis PivotTable item selector.
    /// </summary>
    public sealed class LegacyXlsPivotPageItem {
        /// <summary>
        /// Creates a page-axis PivotTable item selector.
        /// </summary>
        public LegacyXlsPivotPageItem(short fieldIndex, short itemIndex, short objectId) {
            FieldIndex = fieldIndex;
            ItemIndex = itemIndex;
            ObjectId = objectId;
            ItemIndexName = itemIndex == 0x7FFD
                ? "AllItems"
                : $"ItemIndex:{itemIndex}";
        }

        /// <summary>Gets the pivot field index placed on the page axis.</summary>
        public short FieldIndex { get; }

        /// <summary>Gets the pivot item index or the SXPI all-items sentinel.</summary>
        public short ItemIndex { get; }

        /// <summary>Gets a stable name for the selected item index.</summary>
        public string ItemIndexName { get; }

        /// <summary>Gets the object identifier of the page item drop-down arrow.</summary>
        public short ObjectId { get; }
    }
}
