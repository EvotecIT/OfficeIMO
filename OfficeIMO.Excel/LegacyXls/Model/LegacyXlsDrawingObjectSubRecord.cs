namespace OfficeIMO.Excel.LegacyXls.Model {
    /// <summary>
    /// Describes a subrecord discovered inside a BIFF OBJ drawing record.
    /// </summary>
    public sealed class LegacyXlsDrawingObjectSubRecord {
        /// <summary>
        /// Creates OBJ subrecord metadata.
        /// </summary>
        public LegacyXlsDrawingObjectSubRecord(
            ushort subRecordType,
            int offset,
            ushort declaredLength,
            int availableLength) {
            if (offset < 0) {
                throw new ArgumentOutOfRangeException(nameof(offset));
            }

            if (availableLength < 0) {
                throw new ArgumentOutOfRangeException(nameof(availableLength));
            }

            SubRecordType = subRecordType;
            SubRecordTypeKey = $"SubRecordType:0x{subRecordType:X4}";
            SubRecordName = GetSubRecordName(subRecordType);
            Offset = offset;
            DeclaredLength = declaredLength;
            AvailableLength = availableLength;
        }

        /// <summary>Gets the OBJ subrecord type identifier.</summary>
        public ushort SubRecordType { get; }

        /// <summary>Gets a stable hexadecimal key for the OBJ subrecord type.</summary>
        public string SubRecordTypeKey { get; }

        /// <summary>Gets the decoded OBJ subrecord name, or a hexadecimal fallback for unknown types.</summary>
        public string SubRecordName { get; }

        /// <summary>Gets the zero-based byte offset of this subrecord inside the OBJ payload.</summary>
        public int Offset { get; }

        /// <summary>Gets the subrecord payload length declared by the BIFF structure.</summary>
        public ushort DeclaredLength { get; }

        /// <summary>Gets the number of payload bytes available in the containing OBJ record.</summary>
        public int AvailableLength { get; }

        /// <summary>Gets whether the containing OBJ record had enough bytes for the declared subrecord payload.</summary>
        public bool IsComplete => AvailableLength >= DeclaredLength;

        /// <summary>Gets whether this subrecord is an FtLbsData structure whose remaining fields are carried by following Continue records.</summary>
        public bool RequiresContinuation => SubRecordType == 0x0013 && DeclaredLength > 0 && AvailableLength < DeclaredLength;

        /// <summary>Gets whether the subrecord's in-record payload is structurally valid for import metadata.</summary>
        public bool HasSupportedPayload => IsComplete || RequiresContinuation;

        /// <summary>Gets a stable completeness state for import reports.</summary>
        public string CompletionState => RequiresContinuation ? "RequiresContinuation" : IsComplete ? "Complete" : "Truncated";

        private static string GetSubRecordName(ushort subRecordType) {
            return subRecordType switch {
                0x0000 => "FtEnd",
                0x0004 => "FtMacro",
                0x0005 => "FtButton",
                0x0006 => "FtGmo",
                0x0007 => "FtCf",
                0x0008 => "FtPioGrbit",
                0x0009 => "FtPictFmla",
                0x000A => "FtCbls",
                0x000B => "FtRbo",
                0x000C => "FtSbs",
                0x000D => "FtNts",
                0x000E => "ObjLinkFmla",
                0x000F => "FtGboData",
                0x0010 => "FtEdoData",
                0x0011 => "FtRboData",
                0x0012 => "FtCblsData",
                0x0013 => "FtLbsData",
                0x0014 => "ObjLinkFmla",
                0x0015 => "FtCmo",
                0x0016 => "FtReserved",
                _ => $"SubRecordType:0x{subRecordType:X4}"
            };
        }
    }
}
