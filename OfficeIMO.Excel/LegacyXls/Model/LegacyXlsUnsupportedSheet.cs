namespace OfficeIMO.Excel.LegacyXls.Model {
    /// <summary>
    /// Describes a legacy XLS sheet entry that is preserved as import metadata but not projected as a worksheet.
    /// </summary>
    public sealed class LegacyXlsUnsupportedSheet {
        /// <summary>
        /// Creates unsupported legacy sheet metadata.
        /// </summary>
        /// <param name="name">Sheet name from the BoundSheet8 record.</param>
        /// <param name="streamOffset">Byte offset of the sheet substream in the BIFF workbook stream.</param>
        /// <param name="visibility">Legacy sheet visibility flag.</param>
        /// <param name="sheetType">Legacy BoundSheet8 sheet type flag.</param>
        /// <param name="kind">Unsupported sheet category.</param>
        public LegacyXlsUnsupportedSheet(
            string name,
            int streamOffset,
            byte visibility,
            byte sheetType,
            LegacyXlsUnsupportedSheetKind kind) {
            Name = name;
            StreamOffset = streamOffset;
            Visibility = visibility;
            SheetType = sheetType;
            Kind = kind;
        }

        /// <summary>
        /// Gets the sheet name from the BoundSheet8 record.
        /// </summary>
        public string Name { get; }

        /// <summary>
        /// Gets the byte offset of the sheet substream in the BIFF workbook stream.
        /// </summary>
        public int StreamOffset { get; }

        /// <summary>
        /// Gets the legacy visibility flag.
        /// </summary>
        public byte Visibility { get; }

        /// <summary>
        /// Gets the legacy BoundSheet8 sheet type flag.
        /// </summary>
        public byte SheetType { get; }

        /// <summary>
        /// Gets the unsupported sheet category.
        /// </summary>
        public LegacyXlsUnsupportedSheetKind Kind { get; }
    }
}
