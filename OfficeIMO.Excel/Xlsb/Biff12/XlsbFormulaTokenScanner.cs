namespace OfficeIMO.Excel.Xlsb.Biff12;

/// <summary>Traverses complete BIFF12 token boundaries before changing native name or reference operands.</summary>
internal static class XlsbFormulaTokenScanner {
    internal static void Visit(byte[] bytes, int start, int length, Action<int, byte, int> visit) {
        if (start < 0 || length < 0 || start > bytes.Length - length) throw new InvalidDataException("Invalid BIFF12 formula token range.");
        int position = start, end = start + length;
        while (position < end) {
            int tokenOffset = position; byte token = bytes[position++]; int size;
            if (token is >= 0x03 and <= 0x16) size = 0;
            else if (token == 0x01) size = 4;
            else if (token == 0x17) {
                Require(position, 2, end); size = checked(2 + (bytes[position] | bytes[position + 1] << 8) * 2);
            } else if (token == 0x18) {
                Require(position, 1, end);
                size = bytes[position] switch { 0x19 => 13, 0x1D => 5, _ => throw new NotSupportedException("The BIFF12 extended token requires a qualified operand layout before rebinding.") };
            } else if (token == 0x19) {
                Require(position, 3, end); byte flags = bytes[position];
                size = flags == 0x04 ? checked(3 + ((bytes[position + 1] | bytes[position + 2] << 8) + 1) * 2) : 3;
            } else if (token is 0x1C or 0x1D) size = 1;
            else if (token == 0x1E) size = 2;
            else if (token == 0x1F) size = 8;
            else if (token is >= 0x20 and <= 0x7D) size = (token & 0x1F) switch {
                0x00 => 14, 0x01 => 2, 0x02 => 3, 0x03 => 4,
                0x04 or 0x0A or 0x0C => 6, 0x05 or 0x0B or 0x0D => 12,
                0x06 or 0x07 or 0x08 => 6, 0x09 => 2,
                0x19 => 6, 0x1A or 0x1C => 8, 0x1B or 0x1D => 14,
                _ => throw new NotSupportedException("The BIFF12 formula token requires a qualified operand layout before rebinding.")
            };
            else throw new NotSupportedException("The BIFF12 formula token requires a qualified operand layout before rebinding.");
            Require(position, size, end); visit(tokenOffset, token, position); position += size;
        }
    }

    internal static void RewriteNames(byte[] bytes, int start, int length, IReadOnlyDictionary<uint, uint?> names) {
        Visit(bytes, start, length, (tokenOffset, token, operand) => {
            if (token is not (0x23 or 0x43 or 0x63)) return;
            uint old = ReadUInt32(bytes, operand);
            if (!names.TryGetValue(old, out uint? replacement)) throw new InvalidDataException("A native XLSB formula refers to an undeclared defined-name slot.");
            if (replacement.HasValue) WriteUInt32(bytes, operand, replacement.Value);
            else {
                // Err + three parentheses occupies the original five bytes. Keeping width preserves native jumps and mem-expression spans.
                bytes[tokenOffset] = 0x1C; bytes[operand] = 0x17;
                bytes[operand + 1] = bytes[operand + 2] = bytes[operand + 3] = 0x15;
            }
        });
    }

    internal static uint ReadUInt32(byte[] bytes, int offset) {
        Require(offset, 4, bytes.Length); return (uint)(bytes[offset] | bytes[offset + 1] << 8 | bytes[offset + 2] << 16 | bytes[offset + 3] << 24);
    }
    internal static void WriteUInt32(byte[] bytes, int offset, uint value) { Require(offset, 4, bytes.Length); for (int index = 0; index < 4; index++) bytes[offset + index] = (byte)(value >> (index * 8)); }
    private static void Require(int offset, int count, int end) { if (offset < 0 || count < 0 || offset > end - count) throw new InvalidDataException("A native BIFF12 formula token is truncated."); }
}
