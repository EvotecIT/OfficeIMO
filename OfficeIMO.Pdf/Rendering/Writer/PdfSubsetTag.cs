namespace OfficeIMO.Pdf;

/// <summary>Builds the six-letter subset prefix (ISO 32000-1, 9.6.4) from the subset's content.</summary>
/// <remarks>The tag is a pure function of the font bytes, name and embedded glyph set, so repeated saves stay byte-identical.</remarks>
internal static class PdfSubsetTag {
    internal static string Apply(string fontName, IEnumerable<int> glyphs, byte[] fontData) {
        var bytes = new List<byte>(Encoding.UTF8.GetBytes(fontName ?? string.Empty)) { 0 };
        bytes.AddRange(Hash(fontData));
        foreach (int glyph in glyphs) {
            bytes.Add((byte)(glyph >> 24));
            bytes.Add((byte)(glyph >> 16));
            bytes.Add((byte)(glyph >> 8));
            bytes.Add((byte)glyph);
        }
        byte[] hash = Hash(bytes.ToArray());
        var tag = new char[7];
        for (int index = 0; index < 6; index++) tag[index] = (char)('A' + hash[index] % 26);
        tag[6] = '+';
        return new string(tag) + fontName;
    }
    private static byte[] Hash(byte[] bytes) {
#if NET5_0_OR_GREATER
        return System.Security.Cryptography.SHA256.HashData(bytes);
#else
        using var sha256 = System.Security.Cryptography.SHA256.Create();
        return sha256.ComputeHash(bytes);
#endif
    }
}
