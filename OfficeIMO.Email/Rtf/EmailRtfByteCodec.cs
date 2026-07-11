namespace OfficeIMO.Email;

internal static class EmailRtfByteCodec {
    internal static bool TryEncode(string rtf, out byte[] bytes) {
        bytes = new byte[rtf.Length];
        for (int index = 0; index < rtf.Length; index++) {
            if (rtf[index] > byte.MaxValue) {
                bytes = Array.Empty<byte>();
                return false;
            }
            bytes[index] = unchecked((byte)rtf[index]);
        }
        return true;
    }
}
