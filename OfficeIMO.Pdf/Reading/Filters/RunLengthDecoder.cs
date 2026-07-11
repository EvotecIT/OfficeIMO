namespace OfficeIMO.Pdf.Filters;

internal static class RunLengthDecoder {
    public static byte[] Decode(byte[] data) {
        if (!TryDecode(data, int.MaxValue, out byte[] output)) {
            throw new InvalidOperationException("RunLengthDecode output exceeded the configured limit.");
        }

        return output;
    }

    public static bool TryDecode(byte[] data, int maxOutputBytes, out byte[] outputBytes) {
        if (data == null || data.Length == 0) {
            outputBytes = Array.Empty<byte>();
            return true;
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

                if (output.Length + literalCount > maxOutputBytes) {
                    outputBytes = Array.Empty<byte>();
                    return false;
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
            if (output.Length + repeatCount > maxOutputBytes) {
                outputBytes = Array.Empty<byte>();
                return false;
            }

            for (int j = 0; j < repeatCount; j++) {
                output.WriteByte(value);
            }
        }

        outputBytes = output.ToArray();
        return true;
    }
}
