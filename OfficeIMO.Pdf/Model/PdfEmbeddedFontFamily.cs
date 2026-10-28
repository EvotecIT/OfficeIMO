namespace OfficeIMO.Pdf;

/// <summary>
/// Represents a caller-supplied TrueType font family used by generated PDF text.
/// </summary>
public sealed partial class PdfEmbeddedFontFamily {
    private readonly byte[] _regular;
    private readonly byte[]? _bold;
    private readonly byte[]? _italic;
    private readonly byte[]? _boldItalic;
    private readonly bool _regularHasPhysicalSlant;
    private readonly bool _regularHasPhysicalWeight;

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
        _regularHasPhysicalSlant = source._regularHasPhysicalSlant;
        _regularHasPhysicalWeight = source._regularHasPhysicalWeight;
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

    internal PdfEmbeddedFontFamily CreateCoverageStableFallbackSnapshot(string? familyName = null) {
        var snapshot = new PdfEmbeddedFontFamily(familyName ?? FamilyName, _regular, _regularHasPhysicalSlant, _regularHasPhysicalWeight);
        return UsesProgramFontNames ? snapshot.WithProgramFontNames() : snapshot;
    }

    private PdfEmbeddedFontFamily(string familyName, byte[] regular, bool hasPhysicalSlant, bool hasPhysicalWeight = false) {
        FamilyName = familyName;
        _regular = regular;
        _regularHasPhysicalWeight = hasPhysicalWeight;
        _bold = hasPhysicalWeight ? regular : null;
        _regularHasPhysicalSlant = hasPhysicalSlant;
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
