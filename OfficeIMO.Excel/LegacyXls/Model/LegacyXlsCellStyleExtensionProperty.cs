namespace OfficeIMO.Excel.LegacyXls.Model {
    /// <summary>
    /// Describes one XFExt formatting property extension preserved from a BIFF style extension record.
    /// </summary>
    public sealed class LegacyXlsCellStyleExtensionProperty {
        internal LegacyXlsCellStyleExtensionProperty(int index, ushort propertyType, string propertyTypeName, ushort totalByteCount, int dataByteCount) {
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
    }
}
