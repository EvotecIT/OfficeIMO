using System.Globalization;
using System.Security.Cryptography;
using System.Text;

namespace OfficeIMO.Pdf;

#pragma warning disable CA1850 // Static HashData is unavailable on netstandard2.0 and net472.

internal sealed partial class PdfFontResource {
    // Shared immutable decoding evidence for runs without unmapped bytes, indexed by source.
    internal PdfTextDecodingEvidence?[]? DecodingEvidenceBySource;
    public string ResourceName { get; }
    public string BaseFont { get; }
    public string Encoding { get; }
    internal bool HasDeclaredEncoding { get; }
    public string FontSubtype { get; }
    public string? EmbeddedProgramSubtype { get; }
    public bool HasToUnicode { get; }
    public bool IsVerticalWriting { get; }
    public ToUnicodeCMap? CMap { get; }
    internal Lazy<PdfPredefinedCMap>? PredefinedCMap { get; }

    /// <summary>Byte length of the composite code at <paramref name="offset"/>; two bytes unless a predefined CMap says otherwise.</summary>
    internal int GetCompositeCodeLength(byte[] bytes, int offset) => PredefinedCMap?.Value.GetCodeLength(bytes, offset) ?? 2;

    /// <summary>True when the bytes divide into complete composite codes.</summary>
    internal bool HasCompleteCompositeCodes(byte[] bytes, int offset, int count) =>
        PredefinedCMap?.Value.HasCompleteCodes(bytes, offset, count) ?? (count & 1) == 0;
    public IReadOnlyDictionary<int, string>? Differences { get; }
    private readonly byte[]? _embeddedTrueTypeFont;
    public byte[]? EmbeddedTrueTypeFont => _embeddedTrueTypeFont ?? RetainedDrawingResource?._embeddedTrueTypeFont;
    // The family identity hashes the embedded program. A font that receives a drawing program is replaced by a
    // copy with its own identity, so each resource hashes its program only when the family is first read.
    private static readonly string DrawingFontFamilyNotCreated = new string('?', 1);
    private string? _drawingFontFamily;
    private readonly byte[]? _drawingFontFamilyProgramData;
    private readonly PdfDrawingFontProgram? _drawingFontFamilyProgram;
    private string? OwnDrawingFontFamily {
        get {
            string? family = Volatile.Read(ref _drawingFontFamily);
            if (!ReferenceEquals(family, DrawingFontFamilyNotCreated)) return family;
            family = CreateDrawingFontFamily(BaseFont, _drawingFontFamilyProgramData, _drawingFontFamilyProgram);
            Volatile.Write(ref _drawingFontFamily, family);
            return family;
        }
    }
    public string? DrawingFontFamily => OwnDrawingFontFamily ?? RetainedDrawingResource?.OwnDrawingFontFamily;
    public int? FontWeight { get; }
    public int? FontDescriptorFlags { get; }
    internal bool IsSymbolicSubstitute { get; }
    public bool IsBold => PdfFontStyleEvidence.IsBold(BaseFont, FontWeight);
    public bool IsItalic => PdfFontStyleEvidence.IsItalic(BaseFont, FontDescriptorFlags);
    internal PdfType3FontResource? Type3 { get; }
    /// <summary>
    /// Lazily built CID width table, shared by every alias of this font so that each text pass does
    /// not rebuild it from the /W array.
    /// </summary>
    internal sealed class CidWidthSlot {
        internal object? Value;
        internal bool IsBuilt;
    }
    internal CidWidthSlot CidWidths { get; }
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
        int? fontDescriptorFlags = null,
        Lazy<PdfPredefinedCMap>? predefinedCMap = null,
        bool hasDeclaredEncoding = false) {
        ResourceName = resourceName;
        BaseFont = baseFont;
        Encoding = encoding;
        HasDeclaredEncoding = hasDeclaredEncoding;
        FontSubtype = fontSubtype ?? string.Empty;
        EmbeddedProgramSubtype = embeddedProgramSubtype;
        Type3 = type3;
        HasToUnicode = hasToUnicode;
        IsVerticalWriting = isVerticalWriting;
        CMap = cmap;
        PredefinedCMap = predefinedCMap;
        Differences = differences;
        _embeddedTrueTypeFont = embeddedTrueTypeFont;
        _drawingFontFamily = embeddedTrueTypeFont == null ? null : DrawingFontFamilyNotCreated;
        _drawingFontFamilyProgramData = embeddedTrueTypeFont;
        FontWeight = fontWeight;
        FontDescriptorFlags = fontDescriptorFlags;
        IsSymbolicSubstitute = DetectSymbolicSubstitute(baseFont, fontDescriptorFlags);
        CidWidths = new CidWidthSlot();
    }

    private PdfFontResource(string resourceName, PdfFontResource source, PdfDrawingFontProgram? drawingProgram = null) {
        byte[]? embeddedTrueTypeFont = drawingProgram?.Program;
        _drawingProgram = drawingProgram ?? source.DrawingProgram;
        _deferredDrawing = drawingProgram == null ? source._deferredDrawing : null;
        _builtinEncoding = source._builtinEncoding;
        ResourceName = resourceName;
        BaseFont = source.BaseFont;
        Encoding = source.Encoding;
        HasDeclaredEncoding = source.HasDeclaredEncoding;
        FontSubtype = source.FontSubtype;
        EmbeddedProgramSubtype = source.EmbeddedProgramSubtype;
        HasToUnicode = source.HasToUnicode;
        IsVerticalWriting = source.IsVerticalWriting;
        CMap = source.CMap;
        PredefinedCMap = source.PredefinedCMap;
        Differences = source.Differences;
        _embeddedTrueTypeFont = embeddedTrueTypeFont ?? source.EmbeddedTrueTypeFont;
        if (embeddedTrueTypeFont == null) {
            _drawingFontFamily = source.DrawingFontFamily;
        } else {
            _drawingFontFamily = DrawingFontFamilyNotCreated;
            _drawingFontFamilyProgramData = embeddedTrueTypeFont;
            _drawingFontFamilyProgram = drawingProgram;
        }
        FontWeight = source.FontWeight;
        FontDescriptorFlags = source.FontDescriptorFlags;
        IsSymbolicSubstitute = source.IsSymbolicSubstitute;
        Type3 = source.Type3;
        CidWidths = source.CidWidths;
    }

    private static bool DetectSymbolicSubstitute(string baseFont, int? descriptorFlags) {
        if (descriptorFlags is int flags && (flags & 4) != 0) return true;
        int subsetSeparator = baseFont.IndexOf('+');
        int offset = subsetSeparator >= 0 ? subsetSeparator + 1 : 0;
        return HasPrefix("Symbol") || HasPrefix("ZapfDingbats") ||
            HasPrefix("Wingdings") || HasPrefix("Webdings");

        bool HasPrefix(string name) => baseFont.Length - offset >= name.Length &&
            string.Compare(baseFont, offset, name, 0, name.Length, StringComparison.OrdinalIgnoreCase) == 0;
    }

    internal PdfFontResource WithResourceName(string resourceName) =>
        string.Equals(ResourceName, resourceName, StringComparison.Ordinal)
            ? this
            : new PdfFontResource(resourceName, this);

    /// <summary>Synthesized drawing program and mappings when the embedded program needed a Unicode cmap.</summary>
    private readonly PdfDrawingFontProgram? _drawingProgram;
    internal PdfDrawingFontProgram? DrawingProgram => _drawingProgram ?? RetainedDrawingResource?._drawingProgram;

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
            // Resources sharing a program but not an encoding CMap must not share painted-glyph mappings.
            if (drawingProgram.CodeMap != null) {
                byte[] codeMap = System.Text.Encoding.ASCII.GetBytes(drawingProgram.CodeMap.Identity);
                Array.Resize(ref identity, identity.Length + codeMap.Length);
                Buffer.BlockCopy(codeMap, 0, identity, identity.Length - codeMap.Length, codeMap.Length);
            }
            hash = sha256.ComputeHash(identity);
        }
        // Font-family parsing limits each candidate to 256 characters. Keep the digest within
        // that bound so a long PDF BaseFont cannot truncate the identity used to load this face.
        const int digestLength = 24;
        const int maximumFamilyLength = 256;
        string name = string.IsNullOrWhiteSpace(baseFont) ? "PDF embedded font" : baseFont;
        var family = new StringBuilder(name.Substring(0, Math.Min(name.Length, maximumFamilyLength - digestLength - 1)));
        family.Append('-');
        // Drawing family names are parsed as CSS-style family lists. PDF names such as "Arial,Bold"
        // must stay one family, so replace list separators, quotes and escapes.
        family.Replace(',', '-').Replace('"', '-').Replace('\'', '-').Replace('\\', '-');
        for (int i = 0; i < 12; i++) family.Append(hash[i].ToString("x2", CultureInfo.InvariantCulture));
        return family.ToString();
    }

}
