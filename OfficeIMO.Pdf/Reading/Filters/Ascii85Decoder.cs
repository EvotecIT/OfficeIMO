namespace OfficeIMO.Pdf.Filters;

internal static class Ascii85Decoder {
    public static byte[] Decode(byte[] data) {
        if (data == null || data.Length == 0) {
            return Array.Empty<byte>();
        }

        using var output = new MemoryStream();
        uint value = 0;
        int count = 0;

        for (int i = 0; i < data.Length; i++) {
            byte b = data[i];
            if (IsWhitespace(b)) {
                continue;
            }

            if (b == (byte)'~') {
                break;
            }

            if (b == (byte)'z') {
                if (count != 0) {
                    throw new FormatException("Invalid 'z' inside partial ASCII85 group.");
                }

                output.WriteByte(0);
                output.WriteByte(0);
                output.WriteByte(0);
                output.WriteByte(0);
                continue;
            }

            if (b < (byte)'!' || b > (byte)'u') {
                throw new FormatException($"Invalid ASCII85 character '{(char)b}'.");
            }

            value = checked(value * 85 + (uint)(b - (byte)'!'));
            count++;

            if (count == 5) {
                WriteTuple(output, value, 4);
                value = 0;
                count = 0;
            }
        }

        if (count > 1) {
            for (int i = count; i < 5; i++) {
                value = checked(value * 85 + 84);
            }

            WriteTuple(output, value, count - 1);
        }

        return output.ToArray();
    }

    private static bool IsWhitespace(byte value) =>
        value == (byte)' ' || value == (byte)'\t' || value == (byte)'\r' || value == (byte)'\n' || value == (byte)'\f' || value == 0;

    private static void WriteTuple(Stream output, uint value, int bytesToWrite) {
        byte[] tuple = new byte[4];
        tuple[0] = (byte)(value >> 24);
        tuple[1] = (byte)(value >> 16);
        tuple[2] = (byte)(value >> 8);
        tuple[3] = (byte)value;
        output.Write(tuple, 0, bytesToWrite);
    }
}
