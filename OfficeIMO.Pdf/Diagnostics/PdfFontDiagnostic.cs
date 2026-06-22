namespace OfficeIMO.Pdf;

/// <summary>Summary of a PDF font dictionary discovered during diagnostics.</summary>
public sealed class PdfFontDiagnostic {
    internal PdfFontDiagnostic(
        int objectNumber,
        string? subtype,
        string? baseFont,
        string? encoding,
        int? fontDescriptorObjectNumber,
        bool hasEmbeddedFontFile,
        string? embeddedFontFileKind) {
        ObjectNumber = objectNumber;
        Subtype = subtype;
        BaseFont = baseFont;
        Encoding = encoding;
        FontDescriptorObjectNumber = fontDescriptorObjectNumber;
        HasEmbeddedFontFile = hasEmbeddedFontFile;
        EmbeddedFontFileKind = embeddedFontFileKind;
    }

    /// <summary>Font dictionary object number.</summary>
    public int ObjectNumber { get; }

    /// <summary>Font subtype, for example Type1, Type0, TrueType, CIDFontType2, or CIDFontType0.</summary>
    public string? Subtype { get; }

    /// <summary>Base font name from /BaseFont, when present.</summary>
    public string? BaseFont { get; }

    /// <summary>Encoding name from /Encoding, when present as a name.</summary>
    public string? Encoding { get; }

    /// <summary>Referenced font descriptor object number, when present.</summary>
    public int? FontDescriptorObjectNumber { get; }

    /// <summary>True when the font descriptor exposes /FontFile, /FontFile2, or /FontFile3.</summary>
    public bool HasEmbeddedFontFile { get; }

    /// <summary>Embedded font file key, for example FontFile, FontFile2, or FontFile3.</summary>
    public string? EmbeddedFontFileKind { get; }

    /// <summary>True when the font dictionary is one of the built-in base 14 Type1 fonts without an embedded font file.</summary>
    public bool IsStandardBase14Font => string.Equals(Subtype, "Type1", StringComparison.Ordinal) && IsBase14Font(BaseFont);

    /// <summary>True when the font should be reviewed for PDF/A or PDF/UA workflows because no embedded font file was found.</summary>
    public bool RequiresEmbeddingReview => !HasEmbeddedFontFile && !IsStandardBase14Font;

    private static bool IsBase14Font(string? baseFont) {
        if (string.IsNullOrEmpty(baseFont)) {
            return false;
        }

        string normalized = baseFont!.TrimStart('/');
        return string.Equals(normalized, "Courier", StringComparison.Ordinal) ||
            string.Equals(normalized, "Courier-Bold", StringComparison.Ordinal) ||
            string.Equals(normalized, "Courier-Oblique", StringComparison.Ordinal) ||
            string.Equals(normalized, "Courier-BoldOblique", StringComparison.Ordinal) ||
            string.Equals(normalized, "Helvetica", StringComparison.Ordinal) ||
            string.Equals(normalized, "Helvetica-Bold", StringComparison.Ordinal) ||
            string.Equals(normalized, "Helvetica-Oblique", StringComparison.Ordinal) ||
            string.Equals(normalized, "Helvetica-BoldOblique", StringComparison.Ordinal) ||
            string.Equals(normalized, "Times-Roman", StringComparison.Ordinal) ||
            string.Equals(normalized, "Times-Bold", StringComparison.Ordinal) ||
            string.Equals(normalized, "Times-Italic", StringComparison.Ordinal) ||
            string.Equals(normalized, "Times-BoldItalic", StringComparison.Ordinal) ||
            string.Equals(normalized, "Symbol", StringComparison.Ordinal) ||
            string.Equals(normalized, "ZapfDingbats", StringComparison.Ordinal);
    }
}
