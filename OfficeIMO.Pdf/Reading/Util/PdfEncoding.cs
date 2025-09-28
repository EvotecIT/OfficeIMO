namespace OfficeIMO.Pdf;

internal static class PdfEncoding {
    // Latin1 (ISO-8859-1) byte-to-string conversion without depending on Encoding.Latin1
    public static string Latin1GetString(byte[] bytes) {
        var chars = new char[bytes.Length];
        for (int i = 0; i < bytes.Length; i++) chars[i] = (char)bytes[i];
        return new string(chars);
    }

    public static byte[] Latin1GetBytes(string text) {
        System.ArgumentNullException.ThrowIfNull(text);

        var bytes = new byte[text.Length];
        for (int i = 0; i < text.Length; i++) {
            char ch = text[i];
            if (ch > 0xFF) {
                throw new System.ArgumentException($"Character '{ch}' (U+{(int)ch:X4}) is not valid in Latin-1 encoding", nameof(text));
            }
            bytes[i] = (byte)ch;
        }
        return bytes;
    }
}

