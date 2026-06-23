namespace OfficeIMO.Excel.LegacyXls.Model {
    /// <summary>
    /// Preserve-only metadata for BIFF style extension records.
    /// </summary>
    public sealed class LegacyXlsCellStyleExtension {
        internal LegacyXlsCellStyleExtension(
            ushort formatIndex,
            ushort extensionCount,
            int recordOffset,
            ushort recordType,
            int payloadLength)
            : this(
                recordName: "XfExt",
                formatIndex: formatIndex,
                hasFormatIndex: true,
                extensionCount: extensionCount,
                hasExtensionCount: true,
                isBuiltInStyle: null,
                isHidden: null,
                isCustom: null,
                styleCategory: null,
                styleCategoryName: null,
                builtInData: null,
                styleName: null,
                recordOffset: recordOffset,
                recordType: recordType,
                payloadLength: payloadLength) {
        }

        internal LegacyXlsCellStyleExtension(
            string recordName,
            bool isBuiltInStyle,
            bool isHidden,
            bool isCustom,
            byte styleCategory,
            string styleCategoryName,
            ushort builtInData,
            string? styleName,
            int recordOffset,
            ushort recordType,
            int payloadLength)
            : this(
                recordName,
                formatIndex: 0,
                hasFormatIndex: false,
                extensionCount: 0,
                hasExtensionCount: false,
                isBuiltInStyle,
                isHidden,
                isCustom,
                styleCategory,
                styleCategoryName,
                builtInData,
                styleName,
                recordOffset,
                recordType,
                payloadLength) {
        }

        private LegacyXlsCellStyleExtension(
            string recordName,
            ushort formatIndex,
            bool hasFormatIndex,
            ushort extensionCount,
            bool hasExtensionCount,
            bool? isBuiltInStyle,
            bool? isHidden,
            bool? isCustom,
            byte? styleCategory,
            string? styleCategoryName,
            ushort? builtInData,
            string? styleName,
            int recordOffset,
            ushort recordType,
            int payloadLength) {
            if (string.IsNullOrWhiteSpace(recordName)) {
                throw new ArgumentException("The style extension record name cannot be empty.", nameof(recordName));
            }

            if (payloadLength < 0) {
                throw new ArgumentOutOfRangeException(nameof(payloadLength));
            }

            RecordName = recordName;
            FormatIndex = formatIndex;
            HasFormatIndex = hasFormatIndex;
            ExtensionCount = extensionCount;
            HasExtensionCount = hasExtensionCount;
            IsBuiltInStyle = isBuiltInStyle;
            IsHidden = isHidden;
            IsCustom = isCustom;
            StyleCategory = styleCategory;
            StyleCategoryName = styleCategoryName;
            BuiltInData = builtInData;
            StyleName = styleName;
            RecordOffset = recordOffset;
            RecordType = recordType;
            PayloadLength = payloadLength;
        }

        /// <summary>Gets the stable BIFF record name.</summary>
        public string RecordName { get; }

        /// <summary>Gets the XF index extended by the XFExt record.</summary>
        public ushort FormatIndex { get; }

        /// <summary>Gets a value indicating whether <see cref="FormatIndex"/> was declared by this record type.</summary>
        public bool HasFormatIndex { get; }

        /// <summary>Gets the number of formatting extension properties declared by the record.</summary>
        public ushort ExtensionCount { get; }

        /// <summary>Gets a value indicating whether <see cref="ExtensionCount"/> was declared by this record type.</summary>
        public bool HasExtensionCount { get; }

        /// <summary>Gets whether the extended style is a built-in style when declared by the record.</summary>
        public bool? IsBuiltInStyle { get; }

        /// <summary>Gets whether the extended style is hidden from the UI when declared by the record.</summary>
        public bool? IsHidden { get; }

        /// <summary>Gets whether the built-in style has a custom definition when declared by the record.</summary>
        public bool? IsCustom { get; }

        /// <summary>Gets the StyleExt category value when declared by the record.</summary>
        public byte? StyleCategory { get; }

        /// <summary>Gets the StyleExt category name when declared by the record.</summary>
        public string? StyleCategoryName { get; }

        /// <summary>Gets the StyleExt built-in style data value when declared by the record.</summary>
        public ushort? BuiltInData { get; }

        /// <summary>Gets the extended style name when declared by the record.</summary>
        public string? StyleName { get; }

        /// <summary>Gets the byte offset of the source BIFF record.</summary>
        public int RecordOffset { get; }

        /// <summary>Gets the source BIFF record type.</summary>
        public ushort RecordType { get; }

        /// <summary>Gets the source BIFF record payload length in bytes.</summary>
        public int PayloadLength { get; }
    }
}
