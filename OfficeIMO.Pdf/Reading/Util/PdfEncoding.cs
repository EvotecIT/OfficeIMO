namespace OfficeIMO.Pdf;

internal static class PdfEncoding {
    // Latin1 (ISO-8859-1) byte-to-string conversion without depending on Encoding.Latin1
    public static string Latin1GetString(byte[] bytes) {
        var chars = new char[bytes.Length];
        for (int i = 0; i < bytes.Length; i++) chars[i] = (char)bytes[i];
        return new string(chars);
    }

    public static byte[] Latin1GetBytes(string s) {
        var bytes = new byte[s.Length];
        for (int i = 0; i < s.Length; i++) bytes[i] = (byte)(s[i] & 0xFF);
        return bytes;
    }
}

