namespace OfficeIMO.Access {
    internal sealed partial class AccessNativeWriter {
        private static readonly byte[] LetterWeights = { 0x4a, 0x4c, 0x4d, 0x4f, 0x51, 0x53, 0x55, 0x57, 0x59, 0x5b, 0x5c, 0x5e, 0x60, 0x62, 0x64, 0x66, 0x68, 0x69, 0x6b, 0x6d, 0x6f, 0x71, 0x73, 0x75, 0x76, 0x78 };

        // Measure and validate before materializing the single encoded key.
        // Mutation callers reserve it against the same allowance as row decoding.
        private static byte[] Key(Table table, Index index, object?[] row, Action<int>? reserveMetadata = null) {
            int length = 0;
            foreach (int ordinal in index.Columns) {
                object? value = row[ordinal];
                length = checked(length + 1 + (value == null ? 0 : KeyValueLength(table.Columns[ordinal], value)));
                if (length > 510) throw new NotSupportedException("Native creation does not truncate index keys longer than 510 bytes.");
            }
            reserveMetadata?.Invoke(length); byte[] key = new byte[length]; int position = 0;
            foreach (int ordinal in index.Columns) {
                object? value = row[ordinal]; if (value == null) { key[position++] = 0; continue; }
                key[position++] = 0x7f;
                switch (table.Columns[ordinal].Type) {
                    case 2: key[position++] = (byte)value; break;
                    case 3: BigEndian(key, position, (ushort)(unchecked((ushort)(short)value) ^ 0x8000), 2); position += 2; break;
                    case 4: BigEndian(key, position, unchecked((uint)(int)value) ^ 0x80000000); position += 4; break;
                    case 10: WriteTextKey((string)value, key, ref position); break;
                }
            }
            return key;
        }

        private static int KeyValueLength(Column column, object value) {
            switch (column.Type) {
                case 2: return 1;
                case 3: return 2;
                case 4: return 4;
                case 10: return TextKeyLength((string)value);
                default: throw new NotSupportedException("Native index creation currently qualifies Byte, Int16, Int32/AutoNumber and bounded text keys.");
            }
        }

        private static int TextKeyLength(string text) {
            bool directory = text.StartsWith("\u0003", StringComparison.Ordinal);
            int length = directory ? 9 : 2, end = TextKeyEnd(text);
            for (int i = directory ? 1 : 0; i < end; i++) length = checked(length + TextWeightLength(text[i]));
            return length;
        }
        private static int TextKeyEnd(string text) { int end = text.Length; while (end != 0 && text[end - 1] == ' ') end--; return end; }
        private static int TextWeightLength(char c) {
            if (c == '_') return 2;
            if (c >= 'a' && c <= 'z' || c >= 'A' && c <= 'Z' || c >= '0' && c <= '9' || c == ' ') return 1;
            throw new NotSupportedException("Native text keys currently qualify ASCII letters, digits, underscore and spaces. Other indexed text and catalog names require additional collation qualification.");
        }

        private static void WriteTextKey(string value, byte[] key, ref int position) {
            bool directory = value.StartsWith("\u0003", StringComparison.Ordinal);
            // General legacy (1033) weights from independently produced Jet4/ACE12 keys.
            for (int i = directory ? 1 : 0, end = TextKeyEnd(value); i < end; i++) {
                char c = value[i];
                if (c >= 'a' && c <= 'z') key[position++] = LetterWeights[c - 'a'];
                else if (c >= 'A' && c <= 'Z') key[position++] = LetterWeights[c - 'A'];
                else if (c >= '0' && c <= '9') key[position++] = (byte)(0x36 + (c - '0') * 2);
                else if (c == '_') { key[position++] = 0x2b; key[position++] = 3; }
                else key[position++] = 7;
            }
            key[position++] = 1;
            if (directory) {
                // Qualified leading directory control character's unprintable weight.
                key[position++] = 1; key[position++] = 1; key[position++] = 1;
                key[position++] = 0x80; key[position++] = 7; key[position++] = 6; key[position++] = 5;
            }
            key[position++] = 0;
        }
    }
}
