namespace OfficeIMO.Pdf;

/// <summary>Decodes the built-in encodings of the standard Symbol and ZapfDingbats fonts to Unicode.</summary>
internal static class PdfSymbolicFontEncoding {
    private static readonly string[] SymbolMap = BuildMap("SymbolEncoding");
    private static readonly string[] ZapfDingbatsMap = BuildMap("ZapfDingbatsEncoding");

    internal static bool IsSymbolicEncoding(string encoding) =>
        encoding is "SymbolEncoding" or "ZapfDingbatsEncoding";

    internal static string Decode(string encoding, byte[] bytes, int maxOutputCharacters) {
        if (bytes is null || bytes.Length == 0) return string.Empty;
        string[] map = string.Equals(encoding, "ZapfDingbatsEncoding", StringComparison.Ordinal) ? ZapfDingbatsMap : SymbolMap;
        long decodedLength = 0L;
        for (int i = 0; i < bytes.Length; i++) {
            decodedLength += map[bytes[i]].Length;
            if (decodedLength > maxOutputCharacters) {
                throw PdfReadLimitException.Create(PdfReadLimitKind.DecodedTextCharacters, maxOutputCharacters, decodedLength);
            }
        }
        var builder = new System.Text.StringBuilder((int)decodedLength);
        for (int i = 0; i < bytes.Length; i++) builder.Append(map[bytes[i]]);
        return builder.ToString();
    }

    private static string[] BuildMap(string encoding) {
        string?[] names = PdfFontEncodingNames.Get(encoding)!;
        var map = new string[256];
        for (int code = 0; code < map.Length; code++) {
            string? name = names[code];
            map[code] = name == null ? string.Empty
                : PdfFontEncodingNames.TryGetSymbolicUnicode(name, out int scalar) ? char.ConvertFromUtf32(scalar)
                : ResourceResolver.TryDecodeStandardGlyphName(name, out string? text) ? text!
                : string.Empty;
        }
        return map;
    }
}
