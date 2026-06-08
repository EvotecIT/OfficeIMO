using System.Globalization;
using System.Linq;

namespace OfficeIMO.Pdf;

public sealed partial class PdfOptions {
    private static readonly char[] OfficeFontFamilySeparators = { ',', ';' };
    private static readonly char[] OfficeFontFamilyTrimChars = { ' ', '\t', '"', '\'' };

    /// <summary>Default installed sans-serif family candidates used by document converters when they need Unicode-capable generated PDF text.</summary>
    public const string DefaultDocumentFontFamilyFallback = "Arial, Aptos, Calibri, Liberation Sans, DejaVu Sans";

    /// <summary>Default installed monospace family candidates used by document converters for code and preformatted text.</summary>
    public const string DefaultDocumentMonospaceFontFamilyFallback = "Consolas, Courier New, Liberation Mono, DejaVu Sans Mono";

    /// <summary>
    /// Uses an Office-style font family name for generated text, embedding the installed TrueType
    /// family when it is available and otherwise falling back to the nearest PDF standard family.
    /// </summary>
    /// <param name="familyName">Office, CSS, or system font family name such as <c>Aptos</c>, <c>Calibri</c>, or <c>Consolas</c>.</param>
    /// <param name="embedSystemFont">When true, installed TrueType faces are preferred over dependency-free standard PDF font aliases.</param>
    public PdfOptions UseOfficeFontFamily(string? familyName, bool embedSystemFont = true) {
        TryUseOfficeFontFamilyCore(familyName, embedSystemFont, requireEmbeddedFont: false);
        return this;
    }

    /// <summary>
    /// Attempts to use an Office-style font family name and reports whether the generated PDF text font state changed.
    /// </summary>
    /// <param name="familyName">Office, CSS, or system font family name such as <c>Aptos</c>, <c>Calibri</c>, or <c>Consolas</c>.</param>
    /// <param name="embedSystemFont">When true, installed TrueType faces are preferred over dependency-free standard PDF font aliases.</param>
    /// <param name="requireEmbeddedFont">When true, returns true only when the selected default generated font family has an embedded font mapping.</param>
    /// <returns>True when the family changed the generated font state and, when requested, produced an embedded default font mapping.</returns>
    public bool TryUseOfficeFontFamily(string? familyName, bool embedSystemFont = true, bool requireEmbeddedFont = false) {
        return TryUseOfficeFontFamilyCore(familyName, embedSystemFont, requireEmbeddedFont);
    }

    /// <summary>
    /// Attempts to configure the shared document Unicode font fallback family for generated PDF text.
    /// </summary>
    /// <param name="requireEmbeddedFont">When true, returns true only when an installed fallback face was embedded.</param>
    /// <returns>True when the fallback changed the generated font state and, when requested, embedded a default font mapping.</returns>
    public bool TryUseDefaultDocumentFontFallback(bool requireEmbeddedFont = true) {
        return TryUseOfficeFontFamily(DefaultDocumentFontFamilyFallback, embedSystemFont: true, requireEmbeddedFont: requireEmbeddedFont);
    }

    /// <summary>
    /// Attempts to register the shared document monospace fallback family for generated PDF code/preformatted text.
    /// </summary>
    /// <param name="requireEmbeddedFont">When true, returns true only when an installed monospace fallback face was embedded.</param>
    /// <returns>True when the fallback changed the generated font state and, when requested, embedded a monospace font mapping.</returns>
    public bool TryRegisterDefaultDocumentMonospaceFontFallback(bool requireEmbeddedFont = false) {
        string beforeEmbeddedFonts = CaptureEmbeddedFontState();
        RegisterOfficeFontFamily(DefaultDocumentMonospaceFontFamilyFallback, PdfStandardFont.Courier);
        bool changed = !string.Equals(beforeEmbeddedFonts, CaptureEmbeddedFontState(), StringComparison.Ordinal);
        return changed && (!requireEmbeddedFont || HasEmbeddedStandardFontFamily(PdfStandardFont.Courier));
    }

    /// <summary>
    /// Reports whether a generated standard-font family slot currently has an embedded font mapping.
    /// </summary>
    /// <param name="font">Generated PDF font slot or variant.</param>
    /// <returns>True when the normalized font family slot has embedded font data.</returns>
    public bool HasEmbeddedStandardFontFamily(PdfStandardFont font) {
        Guard.StandardFont(font, nameof(font), "PDF embedded font lookup must target one of the supported standard PDF fonts.");
        return _embeddedFonts != null && _embeddedFonts.ContainsKey(PdfStandardFontMapper.GetFontFamily(font));
    }

    private bool TryUseOfficeFontFamilyCore(string? familyName, bool embedSystemFont, bool requireEmbeddedFont) {
        if (string.IsNullOrWhiteSpace(familyName)) {
            return false;
        }

        PdfStandardFont beforeDefault = DefaultFont;
        PdfStandardFont beforeHeader = HeaderFont;
        PdfStandardFont beforeFooter = FooterFont;
        string beforeEmbeddedFonts = CaptureEmbeddedFontState();

        if (PdfStandardFontMapper.TryMapFontFamily(familyName, out PdfStandardFont standardFont)) {
            PdfStandardFont family = PdfStandardFontMapper.GetFontFamily(standardFont);
            RegisterOfficeFontFamily(familyName, family, embedSystemFont);
            DefaultFont = family;
            HeaderFont = family;
            FooterFont = family;
        } else if (embedSystemFont && TryLoadOfficeFontFamily(familyName!, out PdfEmbeddedFontFamily? embeddedFamily) && embeddedFamily != null) {
            RegisterFontFamily(PdfStandardFont.Helvetica, embeddedFamily);
            DefaultFont = PdfStandardFont.Helvetica;
            HeaderFont = PdfStandardFont.Helvetica;
            FooterFont = PdfStandardFont.Helvetica;
        }

        bool changed = beforeDefault != DefaultFont ||
                       beforeHeader != HeaderFont ||
                       beforeFooter != FooterFont ||
                       !string.Equals(beforeEmbeddedFonts, CaptureEmbeddedFontState(), StringComparison.Ordinal);
        return changed && (!requireEmbeddedFont || HasEmbeddedStandardFontFamily(DefaultFont));
    }

    private string CaptureEmbeddedFontState() {
        if (_embeddedFonts == null || _embeddedFonts.Count == 0) {
            return string.Empty;
        }

        return string.Join("|", _embeddedFonts
            .OrderBy(font => font.Key)
            .Select(font => ((int)font.Key).ToString(CultureInfo.InvariantCulture) + ":" + (font.Value.FontName ?? string.Empty) + ":" + font.Value.DataSnapshot.Length.ToString(CultureInfo.InvariantCulture)));
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
