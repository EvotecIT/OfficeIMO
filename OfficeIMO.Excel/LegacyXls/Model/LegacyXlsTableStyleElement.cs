namespace OfficeIMO.Excel.LegacyXls.Model {
    /// <summary>
    /// Represents a BIFF TableStyleElement record.
    /// </summary>
    public sealed class LegacyXlsTableStyleElement {
        /// <summary>
        /// Initializes a new instance of the <see cref="LegacyXlsTableStyleElement"/> class.
        /// </summary>
        public LegacyXlsTableStyleElement(
            uint elementType,
            string elementTypeName,
            uint stripeSize,
            uint differentialFormatIndex,
            ushort headerRecordType,
            ushort headerFlags,
            int recordOffset,
            ushort recordType,
            int payloadLength) {
            ElementType = elementType;
            ElementTypeName = elementTypeName;
            StripeSize = stripeSize;
            DifferentialFormatIndex = differentialFormatIndex;
            HeaderRecordType = headerRecordType;
            HeaderFlags = headerFlags;
            RecordOffset = recordOffset;
            RecordType = recordType;
            PayloadLength = payloadLength;
        }

        /// <summary>Gets the raw table style element type.</summary>
        public uint ElementType { get; }

        /// <summary>Gets the friendly table style element type name.</summary>
        public string ElementTypeName { get; }

        /// <summary>Gets the stripe size for stripe-band elements.</summary>
        public uint StripeSize { get; }

        /// <summary>Gets the differential format index referenced by this element.</summary>
        public uint DifferentialFormatIndex { get; }

        /// <summary>Gets the FRT header record type stored inside the payload.</summary>
        public ushort HeaderRecordType { get; }

        /// <summary>Gets the FRT header flags stored inside the payload.</summary>
        public ushort HeaderFlags { get; }

        /// <summary>Gets the BIFF stream offset of the source record.</summary>
        public int RecordOffset { get; }

        /// <summary>Gets the source BIFF record type.</summary>
        public ushort RecordType { get; }

        /// <summary>Gets the source BIFF payload length.</summary>
        public int PayloadLength { get; }
    }
}
