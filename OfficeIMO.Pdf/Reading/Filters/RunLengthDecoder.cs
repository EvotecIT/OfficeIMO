namespace OfficeIMO.Pdf.Filters;

internal static class RunLengthDecoder {
    public static byte[] Decode(byte[] data) {
        if (data == null || data.Length == 0) {
            return Array.Empty<byte>();
        }

        using var output = new MemoryStream();
        int i = 0;
        while (i < data.Length) {
            byte length = data[i++];
            if (length == 128) {
                break;
            }

            if (length <= 127) {
                int literalCount = length + 1;
                if (i + literalCount > data.Length) {
                    throw new FormatException("RunLengthDecode literal run exceeds input length.");
                }

                output.Write(data, i, literalCount);
                i += literalCount;
                continue;
            }

            int repeatCount = 257 - length;
            if (i >= data.Length) {
                throw new FormatException("RunLengthDecode repeat run is missing a byte.");
            }

            byte value = data[i++];
            for (int j = 0; j < repeatCount; j++) {
                output.WriteByte(value);
            }
        }

        return output.ToArray();
    }
}
