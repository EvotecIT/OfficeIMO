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
    /// <param name="bold">Optional bold TrueType face bytes. When null, the regular face is reused.</param>
    /// <param name="italic">Optional italic TrueType face bytes. When null, the regular face is reused.</param>
    /// <param name="boldItalic">Optional bold-italic TrueType face bytes. When null, the best available supplied face is reused.</param>
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

        FamilyName = familyName.Trim();
        _regular = (byte[])regular.Clone();
        _bold = CloneOptional(bold);
        _italic = CloneOptional(italic);
        _boldItalic = CloneOptional(boldItalic);
    }

    /// <summary>Family name used for generated PDF font resource names.</summary>
    public string FamilyName { get; }

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

    internal PdfEmbeddedFontFamily Clone() =>
        new PdfEmbeddedFontFamily(FamilyName, _regular, _bold, _italic, _boldItalic);

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
            System.IO.File.ReadAllBytes(regularPath),
            string.IsNullOrWhiteSpace(boldPath) ? null : System.IO.File.ReadAllBytes(boldPath!),
            string.IsNullOrWhiteSpace(italicPath) ? null : System.IO.File.ReadAllBytes(italicPath!),
            string.IsNullOrWhiteSpace(boldItalicPath) ? null : System.IO.File.ReadAllBytes(boldItalicPath!));
    }

    private static byte[]? CloneOptional(byte[]? data) =>
        data == null ? null : (byte[])data.Clone();
}
