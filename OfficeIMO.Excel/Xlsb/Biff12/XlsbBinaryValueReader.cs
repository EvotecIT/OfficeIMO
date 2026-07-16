namespace OfficeIMO.Excel.Xlsb.Biff12 {
    /// <summary>Reads bounded BIFF12 scalar encodings shared by workbook and worksheet records.</summary>
    internal static class XlsbBinaryValueReader {
        internal static string? ReadNullableWideString(XlsbBinaryCursor cursor, int maxCharacters) {
            if (cursor == null) throw new ArgumentNullException(nameof(cursor));
            if (maxCharacters < 0) throw new ArgumentOutOfRangeException(nameof(maxCharacters));
            uint count = cursor.ReadUInt32();
            if (count == uint.MaxValue) return null;
            if (count > maxCharacters) {
                throw new InvalidDataException($"The BIFF12 nullable string declares {count} characters, exceeding the configured limit of {maxCharacters} characters.");
            }
            int byteCount;
            try {
                byteCount = checked((int)count * 2);
            } catch (OverflowException exception) {
                throw new InvalidDataException("The BIFF12 nullable string length is too large.", exception);
            }
            return Encoding.Unicode.GetString(cursor.ReadBytes(byteCount));
        }

        internal static bool ReadUInt32Boolean(XlsbBinaryCursor cursor, XlsbRecord record, string fieldName) {
            if (cursor == null) throw new ArgumentNullException(nameof(cursor));
            if (record == null) throw new ArgumentNullException(nameof(record));
            uint value = cursor.ReadUInt32();
            if (value > 1U) {
                throw new InvalidDataException($"The record at offset {record.Offset} contains invalid Boolean {fieldName} value {value}.");
            }
            return value != 0U;
        }
    }
}
