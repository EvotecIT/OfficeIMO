using System.Text;

namespace OfficeIMO.Yaml.Benchmarks.Adapters;

/// <summary>
/// Byte-to-text step for reference libraries that only read text. Detection follows YAML 1.2 section 5.2:
/// byte order marks first, then the null-byte pattern of an ASCII first character, defaulting to UTF-8.
/// The same encoding and byte order mark are used when the text is saved again.
/// </summary>
internal readonly struct YamlTextEncoding {
    private YamlTextEncoding(Encoding encoding, int preambleLength) {
        Encoding = encoding;
        PreambleLength = preambleLength;
    }

    public Encoding Encoding { get; }

    public int PreambleLength { get; }

    public static string Decode(byte[] bytes, out YamlTextEncoding encoding) {
        encoding = Detect(bytes);
        return encoding.Encoding.GetString(bytes, encoding.PreambleLength, bytes.Length - encoding.PreambleLength);
    }

    public byte[] Encode(string text) {
        byte[] preamble = PreambleLength > 0 ? Encoding.GetPreamble() : new byte[0];
        byte[] body = Encoding.GetBytes(text);
        var result = new byte[preamble.Length + body.Length];
        Buffer.BlockCopy(preamble, 0, result, 0, preamble.Length);
        Buffer.BlockCopy(body, 0, result, preamble.Length, body.Length);
        return result;
    }

    private static YamlTextEncoding Detect(byte[] b) {
        var utf32Be = new UTF32Encoding(bigEndian: true, byteOrderMark: true, throwOnInvalidCharacters: true);
        var utf32Le = new UTF32Encoding(bigEndian: false, byteOrderMark: true, throwOnInvalidCharacters: true);
        var utf16Be = new UnicodeEncoding(bigEndian: true, byteOrderMark: true, throwOnInvalidBytes: true);
        var utf16Le = new UnicodeEncoding(bigEndian: false, byteOrderMark: true, throwOnInvalidBytes: true);
        var utf8 = new UTF8Encoding(encoderShouldEmitUTF8Identifier: true, throwOnInvalidBytes: true);
        int n = b.Length;
        if (n >= 4 && b[0] == 0 && b[1] == 0 && b[2] == 0xFE && b[3] == 0xFF) return new YamlTextEncoding(utf32Be, 4);
        if (n >= 4 && b[0] == 0xFF && b[1] == 0xFE && b[2] == 0 && b[3] == 0) return new YamlTextEncoding(utf32Le, 4);
        if (n >= 2 && b[0] == 0xFE && b[1] == 0xFF) return new YamlTextEncoding(utf16Be, 2);
        if (n >= 2 && b[0] == 0xFF && b[1] == 0xFE) return new YamlTextEncoding(utf16Le, 2);
        if (n >= 3 && b[0] == 0xEF && b[1] == 0xBB && b[2] == 0xBF) return new YamlTextEncoding(utf8, 3);
        if (n >= 4 && b[0] == 0 && b[1] == 0 && b[2] == 0) return new YamlTextEncoding(utf32Be, 0);
        if (n >= 4 && b[1] == 0 && b[2] == 0 && b[3] == 0) return new YamlTextEncoding(utf32Le, 0);
        if (n >= 2 && b[0] == 0) return new YamlTextEncoding(utf16Be, 0);
        if (n >= 2 && b[1] == 0) return new YamlTextEncoding(utf16Le, 0);
        return new YamlTextEncoding(utf8, 0);
    }
}
