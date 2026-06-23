namespace OfficeIMO.Excel.LegacyXls.Model {
    /// <summary>
    /// Workbook-level cell style metadata parsed from a BIFF Style record.
    /// </summary>
    public sealed class LegacyXlsCellStyle {
        internal LegacyXlsCellStyle(
            ushort styleFormatIndex,
            bool isBuiltIn,
            byte? builtInStyleId,
            byte? outlineLevel,
            string? name,
            int recordOffset,
            ushort recordType) {
            StyleFormatIndex = styleFormatIndex;
            IsBuiltIn = isBuiltIn;
            BuiltInStyleId = builtInStyleId;
            OutlineLevel = outlineLevel;
            Name = name;
            RecordOffset = recordOffset;
            RecordType = recordType;
        }

        /// <summary>
        /// Gets the zero-based index of the style XF in the workbook XF collection.
        /// </summary>
        public ushort StyleFormatIndex { get; }

        /// <summary>
        /// Gets whether this cell style is one of Excel's built-in styles.
        /// </summary>
        public bool IsBuiltIn { get; }

        /// <summary>
        /// Gets the built-in style identifier when <see cref="IsBuiltIn"/> is true.
        /// </summary>
        public byte? BuiltInStyleId { get; }

        /// <summary>
        /// Gets the outline level for built-in row or column level styles.
        /// </summary>
        public byte? OutlineLevel { get; }

        /// <summary>
        /// Gets the custom style name when <see cref="IsBuiltIn"/> is false.
        /// </summary>
        public string? Name { get; }

        /// <summary>
        /// Gets the byte offset of the source BIFF record.
        /// </summary>
        public int RecordOffset { get; }

        /// <summary>
        /// Gets the source BIFF record type.
        /// </summary>
        public ushort RecordType { get; }
    }
}
