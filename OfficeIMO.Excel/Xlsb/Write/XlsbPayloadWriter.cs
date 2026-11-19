namespace OfficeIMO.Excel.Xlsb.Write {
    /// <summary>Little-endian BIFF12 field writers shared by the record encoders.</summary>
    internal static class XlsbPayloadWriter {
        internal static void WriteUInt16(Stream output, ushort value) {
            output.WriteByte((byte)value);
            output.WriteByte((byte)(value >> 8));
        }

        internal static void WriteUInt32(Stream output, uint value) {
            output.WriteByte((byte)value);
            output.WriteByte((byte)(value >> 8));
            output.WriteByte((byte)(value >> 16));
            output.WriteByte((byte)(value >> 24));
        }

        internal static void WriteInt32(Stream output, int value) => WriteUInt32(output, unchecked((uint)value));

        internal static void WriteDouble(Stream output, double value) {
            ulong bits = unchecked((ulong)BitConverter.DoubleToInt64Bits(value));
            WriteUInt32(output, unchecked((uint)bits));
            WriteUInt32(output, unchecked((uint)(bits >> 32)));
        }

        /// <summary>XLWideString (MS-XLSB 2.5.168): a character count followed by UTF-16 code units.</summary>
        internal static void WriteWideString(Stream output, string value) {
            WriteUInt32(output, checked((uint)value.Length));
            foreach (char character in value) WriteUInt16(output, character);
        }

        /// <summary>XLNullableWideString (MS-XLSB 2.5.167): a null string is the count 0xFFFFFFFF.</summary>
        internal static void WriteNullableWideString(Stream output, string? value) {
            if (value == null) {
                WriteUInt32(output, 0xFFFFFFFFU);
                return;
            }
            WriteWideString(output, value);
        }

        /// <summary>RfX (MS-XLSB 2.5.117) from one-based inclusive coordinates.</summary>
        internal static void WriteRfX(Stream output, int firstRow, int firstColumn, int lastRow, int lastColumn) {
            WriteUInt32(output, checked((uint)(firstRow - 1)));
            WriteUInt32(output, checked((uint)(lastRow - 1)));
            WriteUInt32(output, checked((uint)(firstColumn - 1)));
            WriteUInt32(output, checked((uint)(lastColumn - 1)));
        }

        /// <summary>Writes a space-separated SpreadsheetML reference list as an UncheckedSqRfX (MS-XLSB 2.5.156).</summary>
        internal static void WriteSqRfX(Stream output, string? references, string sheetName, string purpose, int maximumRanges) {
            string[] parts = (references ?? string.Empty).Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
            if (parts.Length == 0 || parts.Length > maximumRanges) {
                throw new NotSupportedException($"Native XLSB generation requires 1 to {maximumRanges} {purpose} ranges on worksheet '{sheetName}'.");
            }
            WriteUInt32(output, (uint)parts.Length);
            foreach (string part in parts) {
                if (!TryParseRange(part, out int firstRow, out int firstColumn, out int lastRow, out int lastColumn)) {
                    throw new NotSupportedException($"Native XLSB generation cannot encode the {purpose} range '{part}' on worksheet '{sheetName}'.");
                }
                WriteRfX(output, firstRow, firstColumn, lastRow, lastColumn);
            }
        }

        /// <summary>Parses an A1 cell or range, with or without <c>$</c>, into one-based inclusive coordinates.</summary>
        internal static bool TryParseRange(string reference, out int firstRow, out int firstColumn, out int lastRow, out int lastColumn) {
            string text = reference.Replace("$", string.Empty);
            if (ExcelA1.TryParseRange(text, out firstRow, out firstColumn, out lastRow, out lastColumn)) {
                if (firstRow > lastRow) (firstRow, lastRow) = (lastRow, firstRow);
                if (firstColumn > lastColumn) (firstColumn, lastColumn) = (lastColumn, firstColumn);
                return firstRow >= 1 && firstColumn >= 1 && lastRow <= 1_048_576 && lastColumn <= 16_384;
            }
            if (ExcelA1.TryParseCellReferenceFast(text, out firstRow, out firstColumn)) {
                lastRow = firstRow;
                lastColumn = firstColumn;
                return firstRow >= 1 && firstColumn >= 1 && firstRow <= 1_048_576 && firstColumn <= 16_384;
            }
            return false;
        }
    }
}
