namespace OfficeIMO.Pdf;

/// <summary>Builds the six-letter subset prefix (ISO 32000-1, 9.6.4) from the subset's content.</summary>
/// <remarks>The tag is a pure function of the font name and the embedded glyph set, so repeated saves stay byte-identical.</remarks>
internal static class PdfSubsetTag {
    internal static string Apply(string fontName, IEnumerable<int> glyphs) {
        var bytes = new List<byte>(Encoding.UTF8.GetBytes(fontName ?? string.Empty)) { 0 };
        foreach (int glyph in glyphs) {
            bytes.Add((byte)(glyph >> 24));
            bytes.Add((byte)(glyph >> 16));
            bytes.Add((byte)(glyph >> 8));
            bytes.Add((byte)glyph);
        }
#if NET5_0_OR_GREATER
        byte[] hash = System.Security.Cryptography.SHA256.HashData(bytes.ToArray());
#else
        byte[] hash;
        using (var sha256 = System.Security.Cryptography.SHA256.Create()) hash = sha256.ComputeHash(bytes.ToArray());
#endif
        var tag = new char[7];
        for (int index = 0; index < 6; index++) tag[index] = (char)('A' + hash[index] % 26);
        tag[6] = '+';
        return new string(tag) + fontName;
    }
}
