namespace OfficeIMO.Excel.LegacyXls.Model {
    /// <summary>
    /// Preserve-only metadata for a BIFF XFExt record that extends an XF cell format.
    /// </summary>
    public sealed class LegacyXlsCellStyleExtension {
        internal LegacyXlsCellStyleExtension(
            ushort formatIndex,
            ushort extensionCount,
            int recordOffset,
            ushort recordType,
            int payloadLength) {
            if (payloadLength < 0) {
                throw new ArgumentOutOfRangeException(nameof(payloadLength));
            }

            FormatIndex = formatIndex;
            ExtensionCount = extensionCount;
            RecordOffset = recordOffset;
            RecordType = recordType;
            PayloadLength = payloadLength;
        }

        /// <summary>Gets the XF index extended by the XFExt record.</summary>
        public ushort FormatIndex { get; }

        /// <summary>Gets the number of formatting extension properties declared by the record.</summary>
        public ushort ExtensionCount { get; }

        /// <summary>Gets the byte offset of the source BIFF record.</summary>
        public int RecordOffset { get; }

        /// <summary>Gets the source BIFF record type.</summary>
        public ushort RecordType { get; }

        /// <summary>Gets the source BIFF record payload length in bytes.</summary>
        public int PayloadLength { get; }
    }
}
