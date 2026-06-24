namespace OfficeIMO.Excel.LegacyXls.Model {
    /// <summary>
    /// Describes a BIFF DVal data-validation collection header discovered during legacy XLS import.
    /// </summary>
    public sealed class LegacyXlsDataValidationCollectionRecord {
        /// <summary>
        /// Creates parsed BIFF DVal collection metadata.
        /// </summary>
        public LegacyXlsDataValidationCollectionRecord(
            string sheetName,
            int recordOffset,
            ushort recordType,
            int payloadLength,
            uint declaredValidationCount) {
            if (payloadLength < 0) {
                throw new ArgumentOutOfRangeException(nameof(payloadLength));
            }

            SheetName = sheetName ?? throw new ArgumentNullException(nameof(sheetName));
            RecordOffset = recordOffset;
            RecordType = recordType;
            PayloadLength = payloadLength;
            DeclaredValidationCount = declaredValidationCount;
        }

        /// <summary>Gets the worksheet name associated with the DVal record.</summary>
        public string SheetName { get; }

        /// <summary>Gets the byte offset of the BIFF record.</summary>
        public int RecordOffset { get; }

        /// <summary>Gets the BIFF record type identifier.</summary>
        public ushort RecordType { get; }

        /// <summary>Gets the BIFF record payload length in bytes.</summary>
        public int PayloadLength { get; }

        /// <summary>Gets the validation-rule count declared by the DVal header.</summary>
        public uint DeclaredValidationCount { get; }
    }
}
