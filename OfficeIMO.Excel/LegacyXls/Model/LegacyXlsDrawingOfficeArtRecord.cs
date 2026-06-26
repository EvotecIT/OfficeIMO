namespace OfficeIMO.Excel.LegacyXls.Model {
    /// <summary>
    /// Describes an OfficeArt record header discovered while traversing an XLS MsoDrawing payload.
    /// </summary>
    public sealed class LegacyXlsDrawingOfficeArtRecord {
        /// <summary>
        /// Creates preserve-only metadata for a nested OfficeArt record.
        /// </summary>
        public LegacyXlsDrawingOfficeArtRecord(ushort recordType, ushort recordInstance, byte recordVersion, uint payloadLength, int depth) {
            if (depth < 0) {
                throw new ArgumentOutOfRangeException(nameof(depth));
            }

            RecordType = recordType;
            RecordTypeKind = LegacyXlsDrawingEscherRecordTypeDecoder.TryGetKind(recordType);
            RecordTypeName = LegacyXlsDrawingEscherRecordTypeDecoder.GetName(recordType);
            RecordInstance = recordInstance;
            RecordVersion = recordVersion;
            PayloadLength = payloadLength;
            Depth = depth;
        }

        /// <summary>Gets the OfficeArt record type identifier.</summary>
        public ushort RecordType { get; }

        /// <summary>Gets the decoded OfficeArt record type, when the identifier is known.</summary>
        public LegacyXlsDrawingEscherRecordType? RecordTypeKind { get; }

        /// <summary>Gets a stable OfficeArt record type display name.</summary>
        public string RecordTypeName { get; }

        /// <summary>Gets the OfficeArt record instance field.</summary>
        public ushort RecordInstance { get; }

        /// <summary>Gets the OfficeArt record version field.</summary>
        public byte RecordVersion { get; }

        /// <summary>Gets the OfficeArt payload length declared by the record header.</summary>
        public uint PayloadLength { get; }

        /// <summary>Gets the traversal depth within nested OfficeArt containers.</summary>
        public int Depth { get; }

        /// <summary>Gets whether this record is an OfficeArt container.</summary>
        public bool IsContainer => RecordVersion == 0x0f;
    }
}
