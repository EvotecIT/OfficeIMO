using System.Globalization;
using System.Security.Cryptography;
using System.Text;

namespace OfficeIMO.Pdf;

#pragma warning disable CA1850 // Static HashData is unavailable on netstandard2.0 and net472.

internal sealed class PdfFontResource {
    public string ResourceName { get; }
    public string BaseFont { get; }
    public string Encoding { get; }
    public string FontSubtype { get; }
    public string? EmbeddedProgramSubtype { get; }
    public bool HasToUnicode { get; }
    public bool IsVerticalWriting { get; }
    public ToUnicodeCMap? CMap { get; }
    public IReadOnlyDictionary<int, string>? Differences { get; }
    public byte[]? EmbeddedTrueTypeFont { get; }
    public string? DrawingFontFamily { get; }
    public int? FontWeight { get; }
    public int? FontDescriptorFlags { get; }
    public bool IsBold => PdfFontStyleEvidence.IsBold(BaseFont, FontWeight);
    public bool IsItalic => PdfFontStyleEvidence.IsItalic(BaseFont, FontDescriptorFlags);
    internal PdfType3FontResource? Type3 { get; }
    public PdfFontResource(
        string resourceName,
        string baseFont,
        string encoding,
        bool hasToUnicode,
        ToUnicodeCMap? cmap = null,
        IReadOnlyDictionary<int, string>? differences = null,
        byte[]? embeddedTrueTypeFont = null,
        string? fontSubtype = null,
        string? embeddedProgramSubtype = null,
        PdfType3FontResource? type3 = null,
        bool isVerticalWriting = false,
        int? fontWeight = null,
        int? fontDescriptorFlags = null) {
        ResourceName = resourceName;
        BaseFont = baseFont;
        Encoding = encoding;
        FontSubtype = fontSubtype ?? string.Empty;
        EmbeddedProgramSubtype = embeddedProgramSubtype;
        Type3 = type3;
        HasToUnicode = hasToUnicode;
        IsVerticalWriting = isVerticalWriting;
        CMap = cmap;
        Differences = differences;
        EmbeddedTrueTypeFont = embeddedTrueTypeFont;
        DrawingFontFamily = CreateDrawingFontFamily(baseFont, embeddedTrueTypeFont);
        FontWeight = fontWeight;
        FontDescriptorFlags = fontDescriptorFlags;
    }

    private PdfFontResource(string resourceName, PdfFontResource source, PdfDrawingFontProgram? drawingProgram = null) {
        byte[]? embeddedTrueTypeFont = drawingProgram?.Program;
        DrawingProgram = drawingProgram ?? source.DrawingProgram;
        ResourceName = resourceName;
        BaseFont = source.BaseFont;
        Encoding = source.Encoding;
        FontSubtype = source.FontSubtype;
        EmbeddedProgramSubtype = source.EmbeddedProgramSubtype;
        HasToUnicode = source.HasToUnicode;
        IsVerticalWriting = source.IsVerticalWriting;
        CMap = source.CMap;
        Differences = source.Differences;
        EmbeddedTrueTypeFont = embeddedTrueTypeFont ?? source.EmbeddedTrueTypeFont;
        DrawingFontFamily = embeddedTrueTypeFont == null
            ? source.DrawingFontFamily
            : CreateDrawingFontFamily(source.BaseFont, embeddedTrueTypeFont, drawingProgram);
        FontWeight = source.FontWeight;
        FontDescriptorFlags = source.FontDescriptorFlags;
        Type3 = source.Type3;
    }

    internal PdfFontResource WithResourceName(string resourceName) =>
        string.Equals(ResourceName, resourceName, StringComparison.Ordinal)
            ? this
            : new PdfFontResource(resourceName, this);

    /// <summary>Synthesized drawing program and mappings when the embedded program needed a Unicode cmap.</summary>
    internal PdfDrawingFontProgram? DrawingProgram { get; }

    /// <summary>Returns this resource with a drawing-ready embedded TrueType program.</summary>
    internal PdfFontResource WithDrawingProgram(PdfDrawingFontProgram drawingProgram) =>
        new PdfFontResource(ResourceName, this, drawingProgram);

    // Embedded programs with the same PDF base name can differ between page and annotation resources.
    private static string? CreateDrawingFontFamily(string baseFont, byte[]? fontData, PdfDrawingFontProgram? drawingProgram = null) {
        if (fontData == null) return null;
        using SHA256 sha256 = SHA256.Create();
        byte[] hash = sha256.ComputeHash(fontData);
        if (drawingProgram != null) {
            // A rebuilt Unicode cmap can omit duplicate or cluster mappings. Keep the full PDF code
            // map in the face identity, including CID entries above the one-byte simple-font range.
            byte[]? cidMap = drawingProgram.CidToGlyphMap;
            var identity = new byte[hash.Length + 1 + (cidMap?.Length ?? 256 * 2)];
            Buffer.BlockCopy(hash, 0, identity, 0, hash.Length);
            identity[hash.Length] = cidMap == null ? (byte)0 : (byte)1;
            if (cidMap != null) {
                Buffer.BlockCopy(cidMap, 0, identity, hash.Length + 1, cidMap.Length);
            } else {
                for (int code = 0; code < 256; code++) {
                    int glyph = drawingProgram.GlyphForCode(code);
                    identity[hash.Length + 1 + code * 2] = (byte)(glyph >> 8);
                    identity[hash.Length + 1 + code * 2 + 1] = (byte)glyph;
                }
            }
            hash = sha256.ComputeHash(identity);
        }
        var family = new StringBuilder(string.IsNullOrWhiteSpace(baseFont) ? "PDF embedded font-" : baseFont + "-");
        // Drawing family names are parsed as CSS-style family lists. PDF names such as "Arial,Bold"
        // must stay one family, so replace list separators, quotes and escapes.
        family.Replace(',', '-').Replace('"', '-').Replace('\'', '-').Replace('\\', '-');
        for (int i = 0; i < 12; i++) family.Append(hash[i].ToString("x2", CultureInfo.InvariantCulture));
        return family.ToString();
    }

}
