namespace OfficeIMO.Excel.LegacyXls.Model {
    /// <summary>
    /// Metadata and projectable formatting facets from BIFF style extension records.
    /// </summary>
    public sealed class LegacyXlsCellStyleExtension {
        internal LegacyXlsCellStyleExtension(
            ushort formatIndex,
            ushort extensionCount,
            int recordOffset,
            ushort recordType,
            int payloadLength,
            IReadOnlyList<LegacyXlsCellStyleExtensionProperty>? properties = null)
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
                associatedStyleFormatIndex: null,
                hasUnparsedStyleProperties: false,
                xfRecordCount: null,
                checksum: null,
                properties,
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
            ushort? associatedStyleFormatIndex,
            bool hasUnparsedStyleProperties,
            IReadOnlyList<LegacyXlsCellStyleExtensionProperty>? properties,
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
                associatedStyleFormatIndex,
                hasUnparsedStyleProperties,
                xfRecordCount: null,
                checksum: null,
                properties,
                recordOffset,
                recordType,
                payloadLength) {
        }

        internal LegacyXlsCellStyleExtension(
            string recordName,
            ushort xfRecordCount,
            uint checksum,
            int recordOffset,
            ushort recordType,
            int payloadLength)
            : this(
                recordName,
                formatIndex: 0,
                hasFormatIndex: false,
                extensionCount: 0,
                hasExtensionCount: false,
                isBuiltInStyle: null,
                isHidden: null,
                isCustom: null,
                styleCategory: null,
                styleCategoryName: null,
                builtInData: null,
                styleName: null,
                associatedStyleFormatIndex: null,
                hasUnparsedStyleProperties: false,
                xfRecordCount,
                checksum,
                properties: null,
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
            ushort? associatedStyleFormatIndex,
            bool hasUnparsedStyleProperties,
            ushort? xfRecordCount,
            uint? checksum,
            IReadOnlyList<LegacyXlsCellStyleExtensionProperty>? properties,
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
            AssociatedStyleFormatIndex = associatedStyleFormatIndex;
            HasUnparsedStyleProperties = hasUnparsedStyleProperties;
            XfRecordCount = xfRecordCount;
            Checksum = checksum;
            Properties = properties == null
                ? Array.Empty<LegacyXlsCellStyleExtensionProperty>()
                : properties.ToArray();
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

        /// <summary>Gets the style XF index from the Style record associated with this StyleExt record, when known.</summary>
        public ushort? AssociatedStyleFormatIndex { get; }

        /// <summary>Gets whether the StyleExt record has trailing formatting properties that are not decoded yet.</summary>
        public bool HasUnparsedStyleProperties { get; }

        /// <summary>Gets the XFCRC-declared number of XF records when declared by the record.</summary>
        public ushort? XfRecordCount { get; }

        /// <summary>Gets the XFCRC checksum when declared by the record.</summary>
        public uint? Checksum { get; }

        /// <summary>Gets the XFExt property extension entries declared by this record.</summary>
        public IReadOnlyList<LegacyXlsCellStyleExtensionProperty> Properties { get; }

        /// <summary>Gets the byte offset of the source BIFF record.</summary>
        public int RecordOffset { get; }

        /// <summary>Gets the source BIFF record type.</summary>
        public ushort RecordType { get; }

        /// <summary>Gets the source BIFF record payload length in bytes.</summary>
        public int PayloadLength { get; }

        /// <summary>Gets whether the parsed style extension is handled without known formatting loss.</summary>
        internal bool IsFullyProjectable {
            get {
                if (IsXfCrc) {
                    return XfRecordCount.HasValue && Checksum.HasValue;
                }

                if (IsXfExt && HasFormatIndex) {
                    return Properties.All(IsProjectableProperty);
                }

                return HasProjectableStyleMetadata;
            }
        }

        /// <summary>Gets whether the style extension carries at least one formatting facet that can be projected.</summary>
        internal bool HasProjectableFormatting => !HasUnparsedStyleProperties
            && Properties.Count > 0
            && Properties.All(IsProjectableProperty)
            && ((IsXfExt && HasFormatIndex) || (IsStyleExt && AssociatedStyleFormatIndex.HasValue));

        /// <summary>Gets whether the extension formatting applies to the specified legacy XF index.</summary>
        internal bool AppliesToFormatIndex(ushort styleIndex) {
            if (IsXfExt && HasFormatIndex) {
                return FormatIndex == styleIndex;
            }

            return IsStyleExt
                && AssociatedStyleFormatIndex.HasValue
                && AssociatedStyleFormatIndex.Value == styleIndex;
        }

        /// <summary>Gets whether the StyleExt record carries workbook style metadata that can be projected.</summary>
        internal bool HasProjectableStyleMetadata => IsStyleExt
            && !HasUnparsedStyleProperties
            && !string.IsNullOrWhiteSpace(StyleName)
            && Properties.All(IsProjectableProperty);

        private bool IsXfExt => string.Equals(RecordName, "XfExt", StringComparison.Ordinal);

        private bool IsStyleExt => string.Equals(RecordName, "StyleExt", StringComparison.Ordinal);

        private bool IsXfCrc => string.Equals(RecordName, "XFCRC", StringComparison.Ordinal);

        private static bool IsProjectableProperty(LegacyXlsCellStyleExtensionProperty property) {
            if (property.UsesStyleXfPropMapping) {
                return IsProjectableStyleXfProp(property);
            }

            if (property.PropertyType == 0x000E) {
                return property.NumericValue == 0x0000
                    || property.NumericValue == 0x0001
                    || property.NumericValue == 0x0002
                    || property.NumericValue == 0x00ff;
            }

            if (property.PropertyType == 0x000F) {
                return property.NumericValue.HasValue;
            }

            if (property.PropertyType == 0x0004
                || property.PropertyType == 0x0005
                || property.PropertyType == 0x0007
                || property.PropertyType == 0x0008
                || property.PropertyType == 0x0009
                || property.PropertyType == 0x000A
                || property.PropertyType == 0x000B
                || property.PropertyType == 0x000D) {
                return property.ColorValue.HasValue
                    && (property.ColorType == 0x0001 || property.ColorType == 0x0002 || property.ColorType == 0x0003);
            }

            return false;
        }

        private static bool IsProjectableStyleXfProp(LegacyXlsCellStyleExtensionProperty property) {
            if (property.PropertyType == 0x0000) {
                return property.NumericValue <= 0x0012;
            }

            if (property.PropertyType == 0x0012) {
                return property.NumericValue <= 15;
            }

            if (property.PropertyType == 0x0025) {
                return property.NumericValue == 0x0000
                    || property.NumericValue == 0x0001
                    || property.NumericValue == 0x0002
                    || property.NumericValue == 0x00ff;
            }

            if (property.PropertyType == 0x0001
                || property.PropertyType == 0x0002
                || property.PropertyType == 0x0005) {
                return HasProjectableColor(property);
            }

            if (property.PropertyType >= 0x0006 && property.PropertyType <= 0x000A) {
                return HasProjectableColor(property)
                    && property.BorderStyle <= 0x000D;
            }

            return false;
        }

        private static bool HasProjectableColor(LegacyXlsCellStyleExtensionProperty property) {
            return property.ColorValue.HasValue
                && (property.ColorType == 0x0001 || property.ColorType == 0x0002 || property.ColorType == 0x0003);
        }
    }
}
