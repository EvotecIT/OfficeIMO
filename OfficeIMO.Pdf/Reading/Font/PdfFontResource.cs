using System.Globalization;
using System.Security.Cryptography;
using System.Text;

namespace OfficeIMO.Pdf;

#pragma warning disable CA1850 // Static HashData is unavailable on netstandard2.0 and net472.

internal sealed class PdfFontResource {
    public string ResourceName { get; }
    public string BaseFont { get; }
    public string Encoding { get; }
    public bool HasToUnicode { get; }
    public ToUnicodeCMap? CMap { get; }
    public IReadOnlyDictionary<int, string>? Differences { get; }
    public byte[]? EmbeddedTrueTypeFont { get; }
    public string? DrawingFontFamily { get; }
    public PdfFontResource(string resourceName, string baseFont, string encoding, bool hasToUnicode, ToUnicodeCMap? cmap = null, IReadOnlyDictionary<int, string>? differences = null, byte[]? embeddedTrueTypeFont = null) {
        ResourceName = resourceName;
        BaseFont = baseFont;
        Encoding = encoding;
        HasToUnicode = hasToUnicode;
        CMap = cmap;
        Differences = differences;
        EmbeddedTrueTypeFont = embeddedTrueTypeFont;
        DrawingFontFamily = CreateDrawingFontFamily(baseFont, embeddedTrueTypeFont);
    }

    private static string? CreateDrawingFontFamily(string baseFont, byte[]? fontData) {
        if (fontData == null || !HasSubsetPrefix(baseFont)) return null;
        using SHA256 sha256 = SHA256.Create();
        byte[] hash = sha256.ComputeHash(fontData);
        var family = new StringBuilder(string.IsNullOrWhiteSpace(baseFont) ? "PDF embedded font-" : baseFont + "-");
        for (int i = 0; i < 12; i++) family.Append(hash[i].ToString("x2", CultureInfo.InvariantCulture));
        return family.ToString();
    }

    private static bool HasSubsetPrefix(string baseFont) {
        if (string.IsNullOrWhiteSpace(baseFont) || baseFont.Length <= 7 || baseFont[6] != '+') return false;
        for (int i = 0; i < 6; i++) {
            char ch = baseFont[i];
            if (ch < 'A' || ch > 'Z') return false;
        }

        return true;
    }
}

