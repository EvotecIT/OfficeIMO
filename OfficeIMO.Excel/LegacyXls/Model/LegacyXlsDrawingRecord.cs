namespace OfficeIMO.Excel.LegacyXls.Model {
    /// <summary>
    /// Describes a preserve-only drawing or object BIFF record discovered during legacy XLS import.
    /// </summary>
    public sealed class LegacyXlsDrawingRecord {
        /// <summary>
        /// Creates drawing BIFF record metadata.
        /// </summary>
        public LegacyXlsDrawingRecord(
            LegacyXlsDrawingRecordKind kind,
            string recordName,
            string? sheetName,
            int recordOffset,
            ushort recordType,
            int payloadLength,
            ushort? objectType = null,
            ushort? objectId = null) {
            if (payloadLength < 0) {
                throw new ArgumentOutOfRangeException(nameof(payloadLength));
            }

            Kind = kind;
            RecordName = recordName ?? throw new ArgumentNullException(nameof(recordName));
            SheetName = sheetName;
            RecordOffset = recordOffset;
            RecordType = recordType;
            PayloadLength = payloadLength;
            ObjectType = objectType;
            ObjectId = objectId;
        }

        /// <summary>Gets the shallow drawing record category.</summary>
        public LegacyXlsDrawingRecordKind Kind { get; }

        /// <summary>Gets the BIFF record name.</summary>
        public string RecordName { get; }

        /// <summary>Gets the worksheet or chart sheet name associated with the record, when known.</summary>
        public string? SheetName { get; }

        /// <summary>Gets the byte offset of the BIFF record.</summary>
        public int RecordOffset { get; }

        /// <summary>Gets the BIFF record type identifier.</summary>
        public ushort RecordType { get; }

        /// <summary>Gets the BIFF record payload length in bytes.</summary>
        public int PayloadLength { get; }

        /// <summary>Gets the decoded OBJ common-object type identifier, when present.</summary>
        public ushort? ObjectType { get; }

        /// <summary>Gets the decoded OBJ object identifier, when present.</summary>
        public ushort? ObjectId { get; }
    }
}
