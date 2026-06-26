namespace OfficeIMO.Excel.LegacyXls.Model {
    /// <summary>
    /// Describes one SXLI PivotTable row or column line item.
    /// </summary>
    public sealed class LegacyXlsPivotLineItem {
        /// <summary>
        /// Creates a PivotTable line item.
        /// </summary>
        public LegacyXlsPivotLineItem(
            short sameAsPreviousCount,
            ushort itemType,
            short entryCount,
            bool multiDataName,
            byte dataIndex,
            bool subtotal,
            bool blockTotal,
            bool grandTotal,
            bool multiDataOnAxis,
            IReadOnlyList<short> entryIndexes) {
            SameAsPreviousCount = sameAsPreviousCount;
            ItemType = itemType;
            ItemTypeKind = TryGetItemTypeKind(itemType);
            ItemTypeName = ItemTypeKind?.ToString() ?? $"LineItemType:{itemType}";
            EntryCount = entryCount;
            MultiDataName = multiDataName;
            DataIndex = dataIndex;
            Subtotal = subtotal;
            BlockTotal = blockTotal;
            GrandTotal = grandTotal;
            MultiDataOnAxis = multiDataOnAxis;
            EntryIndexes = entryIndexes ?? throw new ArgumentNullException(nameof(entryIndexes));
            EntryIndexNames = EntryIndexes.Select(GetEntryIndexName).ToArray();
        }

        /// <summary>Gets the count of leading entries that repeat the previous line item.</summary>
        public short SameAsPreviousCount { get; }

        /// <summary>Gets the raw SXLI item type.</summary>
        public ushort ItemType { get; }

        /// <summary>Gets the decoded SXLI item type, when known.</summary>
        public LegacyXlsPivotLineItemType? ItemTypeKind { get; }

        /// <summary>Gets the SXLI item type name or a stable raw identifier for unknown values.</summary>
        public string ItemTypeName { get; }

        /// <summary>Gets the declared number of line entries.</summary>
        public short EntryCount { get; }

        /// <summary>Gets whether the data field name is used for the total or subtotal.</summary>
        public bool MultiDataName { get; }

        /// <summary>Gets the data item index associated with a subtotal.</summary>
        public byte DataIndex { get; }

        /// <summary>Gets whether the line item represents a subtotal.</summary>
        public bool Subtotal { get; }

        /// <summary>Gets whether the line item represents a block total.</summary>
        public bool BlockTotal { get; }

        /// <summary>Gets whether the line item represents a grand total.</summary>
        public bool GrandTotal { get; }

        /// <summary>Gets whether a line entry represents a data item index.</summary>
        public bool MultiDataOnAxis { get; }

        /// <summary>Gets the raw pivot line entry indexes.</summary>
        public IReadOnlyList<short> EntryIndexes { get; }

        /// <summary>Gets stable names for the pivot line entry indexes.</summary>
        public IReadOnlyList<string> EntryIndexNames { get; }

        private static LegacyXlsPivotLineItemType? TryGetItemTypeKind(ushort value) {
            return value <= 14 ? (LegacyXlsPivotLineItemType)value : null;
        }

        private static string GetEntryIndexName(short value) {
            return value == 0x7FFF
                ? "BlankEntry"
                : $"EntryIndex:{value}";
        }
    }
}
