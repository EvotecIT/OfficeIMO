namespace OfficeIMO.Pdf;

/// <summary>
/// Represents a caller-supplied TrueType font family used by generated PDF text.
/// </summary>
public sealed partial class PdfEmbeddedFontFamily {
    private readonly byte[] _regular;
    private readonly byte[]? _bold;
    private readonly byte[]? _italic;
    private readonly byte[]? _boldItalic;

    /// <summary>Creates a reusable TrueType font family from font file bytes.</summary>
    /// <param name="familyName">Family name used for generated PDF font resource names.</param>
    /// <param name="regular">Regular TrueType face bytes.</param>
    /// <param name="bold">Optional bold TrueType face bytes. When null, the selected regular face is emboldened while painting.</param>
    /// <param name="italic">Optional italic TrueType face bytes. When null, generated italic text uses a slanted regular face.</param>
    /// <param name="boldItalic">Optional bold-italic TrueType face bytes. When null, generated text uses the best supplied face and simulates the missing slant when needed.</param>
    public PdfEmbeddedFontFamily(
        string familyName,
        byte[] regular,
        byte[]? bold = null,
        byte[]? italic = null,
        byte[]? boldItalic = null) {
        Guard.NotNullOrWhiteSpace(familyName, nameof(familyName));
        Guard.NotNullOrEmpty(regular, nameof(regular));
        if (bold != null) {
            Guard.NotNullOrEmpty(bold, nameof(bold));
        }

        if (italic != null) {
            Guard.NotNullOrEmpty(italic, nameof(italic));
        }

        if (boldItalic != null) {
            Guard.NotNullOrEmpty(boldItalic, nameof(boldItalic));
        }

        PdfFontInput.EnsureWithinLimit(regular.LongLength);
        if (bold != null) PdfFontInput.EnsureWithinLimit(bold.LongLength);
        if (italic != null) PdfFontInput.EnsureWithinLimit(italic.LongLength);
        if (boldItalic != null) PdfFontInput.EnsureWithinLimit(boldItalic.LongLength);

        FamilyName = familyName.Trim();
        _regular = SnapshotFace(regular, familyName);
        _bold = bold == null ? null : SnapshotFace(bold, familyName);
        _italic = italic == null ? null : SnapshotFace(italic, familyName);
        _boldItalic = boldItalic == null ? null : SnapshotFace(boldItalic, familyName);
    }

    // Faces read from disk by the library are private to it, so they need no defensive snapshot copy.
    private PdfEmbeddedFontFamily(string familyName, byte[] regular, byte[]? bold, byte[]? italic, byte[]? boldItalic, bool facesOwned) {
        Guard.NotNullOrWhiteSpace(familyName, nameof(familyName));
        Guard.NotNullOrEmpty(regular, nameof(regular));
        PdfFontInput.EnsureWithinLimit(regular.LongLength);
        if (bold != null) PdfFontInput.EnsureWithinLimit(bold.LongLength);
        if (italic != null) PdfFontInput.EnsureWithinLimit(italic.LongLength);
        if (boldItalic != null) PdfFontInput.EnsureWithinLimit(boldItalic.LongLength);
        FamilyName = familyName.Trim();
        _regular = regular;
        _bold = bold;
        _italic = italic;
        _boldItalic = boldItalic;
    }

    /// <summary>Creates a family from face bytes that no caller can reach or change.</summary>
    internal static PdfEmbeddedFontFamily FromOwnedFaces(string familyName, byte[] regular, byte[]? bold, byte[]? italic, byte[]? boldItalic) =>
        new PdfEmbeddedFontFamily(familyName, regular, bold, italic, boldItalic, facesOwned: true);

    /// <summary>Family name used for generated PDF font resource names.</summary>
    public string FamilyName { get; }

    // Converter resource identities can include a style or content fingerprint. Those
    // remain lookup keys; the embedded program supplies the reader-visible font name.
    internal bool UsesProgramFontNames { get; }

    internal PdfEmbeddedFontFamily WithProgramFontNames() =>
        UsesProgramFontNames ? this : new PdfEmbeddedFontFamily(this);

    private PdfEmbeddedFontFamily(PdfEmbeddedFontFamily source) {
        FamilyName = source.FamilyName;
        _regular = source._regular;
        _bold = source._bold;
        _italic = source._italic;
        _boldItalic = source._boldItalic;
        UsesProgramFontNames = true;
    }

    /// <summary>Regular TrueType face bytes.</summary>
    public byte[] Regular => (byte[])_regular.Clone();

    /// <summary>Optional bold TrueType face bytes.</summary>
    public byte[]? Bold => CloneOptional(_bold);

    /// <summary>Optional italic TrueType face bytes.</summary>
    public byte[]? Italic => CloneOptional(_italic);

    /// <summary>Optional bold-italic TrueType face bytes.</summary>
    public byte[]? BoldItalic => CloneOptional(_boldItalic);

    internal byte[] RegularSnapshot => _regular;

    internal byte[]? BoldSnapshot => _bold;

    internal byte[]? ItalicSnapshot => _italic;

    internal byte[]? BoldItalicSnapshot => _boldItalic;

    internal PdfEmbeddedFontFamily Clone() => this;

    // Fallback text is planned against the regular face. Styled runs use the family's own faces
    // where those faces have the glyphs; a face whose licence forbids embedding, or that cannot be
    // parsed, is left out so that style is synthesized from the regular face.
    internal PdfEmbeddedFontFamily CreateFallbackSnapshot() {
        var snapshot = new PdfEmbeddedFontFamily(FamilyName, _regular,
            EmbeddableFace(_bold), EmbeddableFace(_italic), EmbeddableFace(_boldItalic), facesOwned: true);
        return UsesProgramFontNames ? snapshot.WithProgramFontNames() : snapshot;
    }

    /// <summary>Copies a fallback snapshot under another resource name.</summary>
    internal PdfEmbeddedFontFamily WithFallbackFamilyName(string familyName) {
        var snapshot = new PdfEmbeddedFontFamily(familyName, _regular, _bold, _italic, _boldItalic, facesOwned: true);
        return UsesProgramFontNames ? snapshot.WithProgramFontNames() : snapshot;
    }

    private byte[]? EmbeddableFace(byte[]? face) =>
        face != null && (ReferenceEquals(face, _regular) || HasUsableProgram(face)) ? face : null;

    /// <summary>
    /// Returns a test that tells whether a styled face paints every scalar of a text that the regular face paints,
    /// or null when the styled face is the regular face itself.
    /// </summary>
    internal static Func<string, bool>? CreateStyledFaceCoverageTest(byte[] styledFace, byte[] regularFace) {
        if (ReferenceEquals(styledFace, regularFace)) return null;
        Func<int, bool> styled = CreateGlyphLookup(styledFace);
        Func<int, bool> regular = CreateGlyphLookup(regularFace);
        return text => {
            for (int index = 0; index < text.Length; index++) {
                int scalar = text[index];
                if (char.IsHighSurrogate(text[index]) && index + 1 < text.Length && char.IsLowSurrogate(text[index + 1])) {
                    scalar = char.ConvertToUtf32(text[index], text[index + 1]);
                    index++;
                }
                if (regular(scalar) && !styled(scalar)) return false;
            }
            return true;
        };
    }

    private static Func<int, bool> CreateGlyphLookup(byte[] face) {
        if (IsOpenTypeCffData(face)) {
            PdfOpenTypeCffFontProgram cff = PdfFontProgramCache.GetOpenTypeCff(face, fontNameOverride: null);
            return scalar => cff.TryGetGlyphId(scalar, out int glyphId) && glyphId > 0;
        }
        PdfTrueTypeFontProgram trueType = PdfFontProgramCache.GetTrueType(face, fontNameOverride: null);
        return scalar => trueType.TryGetGlyphId(scalar, out int glyphId) && glyphId > 0;
    }

    private PdfEmbeddedFontFamily(string familyName, byte[] regular, bool hasPhysicalSlant, bool hasPhysicalWeight = false) {
        FamilyName = familyName;
        _regular = regular;
        _bold = hasPhysicalWeight ? regular : null;
        _italic = hasPhysicalSlant ? regular : null;
        _boldItalic = hasPhysicalSlant && hasPhysicalWeight ? regular : null;
    }

    /// <summary>Creates a reusable TrueType font family from font files on disk.</summary>
    public static PdfEmbeddedFontFamily FromFiles(
        string familyName,
        string regularPath,
        string? boldPath = null,
        string? italicPath = null,
        string? boldItalicPath = null) {
        Guard.NotNullOrWhiteSpace(regularPath, nameof(regularPath));
        return new PdfEmbeddedFontFamily(
            familyName,
            PdfFontInput.ReadFile(regularPath),
            string.IsNullOrWhiteSpace(boldPath) ? null : PdfFontInput.ReadFile(boldPath!),
            string.IsNullOrWhiteSpace(italicPath) ? null : PdfFontInput.ReadFile(italicPath!),
            string.IsNullOrWhiteSpace(boldItalicPath) ? null : PdfFontInput.ReadFile(boldItalicPath!));
    }

    private static byte[]? CloneOptional(byte[]? data) =>
        data == null ? null : (byte[])data.Clone();
}
