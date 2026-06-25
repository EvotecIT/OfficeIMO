namespace OfficeIMO.Excel.LegacyXls.Model {
    /// <summary>
    /// Describes a preserve-only extended metadata BIFF record found inside a worksheet or chart-sheet substream.
    /// </summary>
    public sealed class LegacyXlsWorksheetFutureMetadataRecord {
        /// <summary>
        /// Creates worksheet future metadata record provenance.
        /// </summary>
        public LegacyXlsWorksheetFutureMetadataRecord(
            LegacyXlsWorkbookMetadataKind kind,
            int recordOffset,
            ushort recordType,
            int payloadLength,
            ushort? headerRecordType,
            ushort? headerFlags) {
            if (payloadLength < 0) {
                throw new ArgumentOutOfRangeException(nameof(payloadLength));
            }

            Kind = kind;
            RecordOffset = recordOffset;
            RecordType = recordType;
            PayloadLength = payloadLength;
            HeaderRecordType = headerRecordType;
            HeaderFlags = headerFlags;
        }

        /// <summary>Gets the decoded metadata kind.</summary>
        public LegacyXlsWorkbookMetadataKind Kind { get; }

        /// <summary>Gets the byte offset of the source BIFF record.</summary>
        public int RecordOffset { get; }

        /// <summary>Gets the BIFF record type identifier.</summary>
        public ushort RecordType { get; }

        /// <summary>Gets the BIFF record payload length in bytes.</summary>
        public int PayloadLength { get; }

        /// <summary>Gets the future-record header record type, when the payload is long enough to expose it.</summary>
        public ushort? HeaderRecordType { get; }

        /// <summary>Gets the future-record option flags, when the payload is long enough to expose them.</summary>
        public ushort? HeaderFlags { get; }

        /// <summary>Gets whether the payload starts with a matching future-record header.</summary>
        public bool HasMatchingFutureRecordHeader => HeaderRecordType == RecordType && PayloadLength >= 12;

        /// <summary>Gets a compact header classification for import reports.</summary>
        public string HeaderState {
            get {
                if (!HeaderRecordType.HasValue) {
                    return "MissingHeader";
                }

                return HasMatchingFutureRecordHeader
                    ? "MatchingFutureHeader"
                    : HeaderRecordType.Value == RecordType
                        ? "ShortFutureHeader"
                        : "RawPayload";
            }
        }

        /// <summary>Gets the number of bytes after the decoded future-record header, or the raw payload length when no matching header is present.</summary>
        public int BodyByteCount => HasMatchingFutureRecordHeader ? PayloadLength - 12 : PayloadLength;
    }
}
