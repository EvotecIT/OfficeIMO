namespace OfficeIMO.Pdf;

public sealed partial class PdfOptions {
    private static readonly char[] OfficeFontFamilySeparators = { ',', ';' };
    private static readonly char[] OfficeFontFamilyTrimChars = { ' ', '\t', '"', '\'' };
    private PdfEmbeddedFontFallbackSet? _embeddedFontFallbacks;

    /// <summary>
    /// Embedded font fallback set used to split generated rich text runs that cannot be written by their selected font.
    /// </summary>
    public PdfEmbeddedFontFallbackSet? EmbeddedFontFallbacks {
        get => _embeddedFontFallbacks?.Clone();
        set {
            _embeddedFontFallbacks = value?.Clone();
            _embeddedFontFallbacks?.RegisterFonts(this);
        }
    }

    internal PdfEmbeddedFontFallbackSet? EmbeddedFontFallbacksSnapshot => _embeddedFontFallbacks?.Clone();

    /// <summary>
    /// Uses an Office-style font family name for generated text, embedding the installed TrueType
    /// family when it is available and otherwise falling back to the nearest PDF standard family.
    /// </summary>
    /// <param name="familyName">Office, CSS, or system font family name such as <c>Aptos</c>, <c>Calibri</c>, or <c>Consolas</c>.</param>
    /// <param name="embedSystemFont">When true, installed TrueType faces are preferred over dependency-free standard PDF font aliases.</param>
    public PdfOptions UseOfficeFontFamily(string? familyName, bool embedSystemFont = true) {
        if (string.IsNullOrWhiteSpace(familyName)) {
            return this;
        }

        if (PdfStandardFontMapper.TryMapFontFamily(familyName, out PdfStandardFont standardFont)) {
            PdfStandardFont family = PdfStandardFontMapper.GetFontFamily(standardFont);
            RegisterOfficeFontFamily(familyName, family, embedSystemFont);
            DefaultFont = family;
            HeaderFont = family;
            FooterFont = family;
            return this;
        }

        if (embedSystemFont && TryLoadOfficeFontFamily(familyName!, out PdfEmbeddedFontFamily? embeddedFamily) && embeddedFamily != null) {
            RegisterFontFamily(PdfStandardFont.Helvetica, embeddedFamily);
            DefaultFont = PdfStandardFont.Helvetica;
            HeaderFont = PdfStandardFont.Helvetica;
            FooterFont = PdfStandardFont.Helvetica;
        }

        return this;
    }

    /// <summary>
    /// Registers an Office-style font family for one semantic PDF font family slot without changing
    /// the document default font. This is used by converters for run-level and cell-level fonts.
    /// </summary>
    /// <param name="familyName">Office, CSS, or system font family name such as <c>Aptos</c>, <c>Georgia</c>, or <c>Consolas</c>.</param>
    /// <param name="baseFontFamily">The semantic standard family slot to back: Helvetica, Times-Roman, or Courier.</param>
    /// <param name="embedSystemFont">When true, installed TrueType faces are embedded into the selected semantic slot when available.</param>
    public PdfOptions RegisterOfficeFontFamily(string? familyName, PdfStandardFont baseFontFamily, bool embedSystemFont = true) {
        if (string.IsNullOrWhiteSpace(familyName)) {
            return this;
        }

        PdfStandardFont normalizedFamily = PdfStandardFontMapper.GetFontFamily(baseFontFamily);
        if (embedSystemFont && TryLoadOfficeFontFamily(familyName!, out PdfEmbeddedFontFamily? embeddedFamily) && embeddedFamily != null) {
            RegisterFontFamily(normalizedFamily, embeddedFamily);
        }

        return this;
    }

    /// <summary>
    /// Registers a caller-supplied TrueType font family for one semantic PDF font family slot without
    /// changing the document default font. Use this for private, licensed, or packaged fonts that
    /// should back a specific generated Helvetica, Times, or Courier family.
    /// </summary>
    /// <param name="baseFontFamily">The semantic standard family slot to back: Helvetica, Times-Roman, or Courier.</param>
    /// <param name="fontFamily">Reusable TrueType font family to embed into that semantic slot.</param>
    public PdfOptions RegisterFontFamily(PdfStandardFont baseFontFamily, PdfEmbeddedFontFamily fontFamily) {
        Guard.NotNull(fontFamily, nameof(fontFamily));
        PdfStandardFont normalizedFamily = PdfStandardFontMapper.GetFontFamily(baseFontFamily);
        PdfEmbeddedFontFamily snapshot = fontFamily.Clone();

        EmbedStandardFont(normalizedFamily, snapshot.RegularSnapshot, BuildFontFamilyFaceName(snapshot.FamilyName, "Regular"));
        EmbedStandardFont(PdfStandardFontMapper.GetStyledFont(normalizedFamily, bold: true, italic: false), snapshot.BoldSnapshot ?? snapshot.RegularSnapshot, BuildFontFamilyFaceName(snapshot.FamilyName, "Bold"));
        EmbedStandardFont(PdfStandardFontMapper.GetStyledFont(normalizedFamily, bold: false, italic: true), snapshot.ItalicSnapshot ?? snapshot.RegularSnapshot, BuildFontFamilyFaceName(snapshot.FamilyName, "Italic"));
        EmbedStandardFont(PdfStandardFontMapper.GetStyledFont(normalizedFamily, bold: true, italic: true), snapshot.BoldItalicSnapshot ?? snapshot.BoldSnapshot ?? snapshot.ItalicSnapshot ?? snapshot.RegularSnapshot, BuildFontFamilyFaceName(snapshot.FamilyName, "BoldItalic"));
        return this;
    }

    /// <summary>
    /// Registers a planned embedded-font fallback set into its generated standard-font family slots.
    /// </summary>
    /// <param name="fallbackSet">Fallback set that pairs prioritized embedded font candidates with generated font slots.</param>
    public PdfOptions RegisterEmbeddedFontFallbacks(PdfEmbeddedFontFallbackSet fallbackSet) {
        Guard.NotNull(fallbackSet, nameof(fallbackSet));
        _embeddedFontFallbacks = fallbackSet.Clone();
        _embeddedFontFallbacks.RegisterFonts(this);
        return this;
    }

    private static bool TryLoadOfficeFontFamily(string familyName, out PdfEmbeddedFontFamily? embeddedFamily) {
        foreach (string candidate in EnumerateOfficeFontFamilyCandidates(familyName)) {
            if (PdfEmbeddedFontFamily.TryFromSystem(candidate, out embeddedFamily) && embeddedFamily != null) {
                return true;
            }
        }

        embeddedFamily = null;
        return false;
    }

    private static System.Collections.Generic.IEnumerable<string> EnumerateOfficeFontFamilyCandidates(string familyName) {
        foreach (string value in familyName.Split(OfficeFontFamilySeparators)) {
            string candidate = value.Trim(OfficeFontFamilyTrimChars);
            if (!string.IsNullOrWhiteSpace(candidate)) {
                yield return candidate;
            }
        }
    }
}
