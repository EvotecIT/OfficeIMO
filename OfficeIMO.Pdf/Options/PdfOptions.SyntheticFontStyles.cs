namespace OfficeIMO.Pdf;

public sealed partial class PdfCreateOptions {
    /// <summary>Whether the requested weight has no supplied physical face and needs native PDF stroke paint.</summary>
    internal bool NeedsSyntheticBold(PdfStandardFont font, PdfNamedFontFace? namedFont = null) {
        if (namedFont.HasValue) {
            PdfNamedFontFace face = namedFont.Value;
            if (!face.Bold || _namedFontFamilies == null ||
                !_namedFontFamilies.TryGetValue(face.FamilyKey, out PdfEmbeddedFontFamily? family)) return false;
            // Match TryGetNamedFontData's preference: bold-italic, bold, italic, regular.
            return face.Italic
                ? family.BoldItalicSnapshot == null && family.BoldSnapshot == null
                : family.BoldSnapshot == null;
        }
        return _embeddedFonts != null && _embeddedFonts.TryGetValue(font, out PdfEmbeddedFont? embedded) &&
            embedded.SyntheticBold;
    }
}
