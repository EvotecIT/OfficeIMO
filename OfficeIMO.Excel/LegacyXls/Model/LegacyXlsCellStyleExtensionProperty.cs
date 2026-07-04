namespace OfficeIMO.Excel.LegacyXls.Model {
    /// <summary>
    /// Describes one XFExt formatting property extension preserved from a BIFF style extension record.
    /// </summary>
    public sealed class LegacyXlsCellStyleExtensionProperty {
        internal LegacyXlsCellStyleExtensionProperty(
            int index,
            ushort propertyType,
            string propertyTypeName,
            ushort totalByteCount,
            int dataByteCount,
            ushort? numericValue = null,
            string? numericValueName = null,
            ushort? colorType = null,
            string? colorTypeName = null,
            short? colorTintShade = null,
            uint? colorValue = null,
            bool usesStyleXfPropMapping = false,
            ushort? borderStyle = null,
            string? borderStyleName = null) {
            if (index < 0) {
                throw new ArgumentOutOfRangeException(nameof(index));
            }

            if (dataByteCount < 0) {
                throw new ArgumentOutOfRangeException(nameof(dataByteCount));
            }

            Index = index;
            PropertyType = propertyType;
            PropertyTypeName = propertyTypeName ?? throw new ArgumentNullException(nameof(propertyTypeName));
            TotalByteCount = totalByteCount;
            DataByteCount = dataByteCount;
            NumericValue = numericValue;
            NumericValueName = string.IsNullOrWhiteSpace(numericValueName) ? null : numericValueName;
            ColorType = colorType;
            ColorTypeName = string.IsNullOrWhiteSpace(colorTypeName) ? null : colorTypeName;
            ColorTintShade = colorTintShade;
            ColorValue = colorValue;
            UsesStyleXfPropMapping = usesStyleXfPropMapping;
            BorderStyle = borderStyle;
            BorderStyleName = string.IsNullOrWhiteSpace(borderStyleName) ? null : borderStyleName;
        }

        /// <summary>Gets the zero-based property index within the XFExt record.</summary>
        public int Index { get; }

        /// <summary>Gets the raw ExtProp type identifier.</summary>
        public ushort PropertyType { get; }

        /// <summary>Gets the decoded ExtProp type name, when known.</summary>
        public string PropertyTypeName { get; }

        /// <summary>Gets the total ExtProp byte count declared by the property.</summary>
        public ushort TotalByteCount { get; }

        /// <summary>Gets the byte count of the ExtProp data payload after the type and size fields.</summary>
        public int DataByteCount { get; }

        /// <summary>Gets a decoded 2-byte numeric value for simple ExtProp payloads, when present.</summary>
        public ushort? NumericValue { get; }

        /// <summary>Gets a decoded name for the simple numeric payload, when known.</summary>
        public string? NumericValueName { get; }

        /// <summary>Gets the FullColorExt color storage type, when this property carries a color payload.</summary>
        public ushort? ColorType { get; }

        /// <summary>Gets the decoded FullColorExt color storage type name, when known.</summary>
        public string? ColorTypeName { get; }

        /// <summary>Gets the FullColorExt tint/shade adjustment, when this property carries a color payload.</summary>
        public short? ColorTintShade { get; }

        /// <summary>Gets the raw FullColorExt color value, when this property carries a color payload.</summary>
        public uint? ColorValue { get; }

        /// <summary>Gets the raw FullColorExt color value formatted as hexadecimal, when present.</summary>
        public string? ColorValueHex => ColorValue.HasValue ? $"0x{ColorValue.Value:X8}" : null;

        /// <summary>Gets whether the property type follows StyleExt XFProps numbering instead of XFExt numbering.</summary>
        internal bool UsesStyleXfPropMapping { get; }

        /// <summary>Gets the XFPropBorder border style value, when present.</summary>
        public ushort? BorderStyle { get; }

        /// <summary>Gets the decoded XFPropBorder border style name, when known.</summary>
        public string? BorderStyleName { get; }
    }
}
