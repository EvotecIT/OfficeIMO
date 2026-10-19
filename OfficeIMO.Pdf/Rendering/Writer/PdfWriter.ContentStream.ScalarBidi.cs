using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal sealed partial class ContentStreamBuilder {
    // Flat LTR/RTL letter runs have an unambiguous ToUnicode representation when each
    // direction has its own show operator. This avoids broad ActualText being reordered
    // twice by common readers. Nested/numeric or non-letter contexts retain ActualText:
    // visual fragment geometry alone cannot recover their original logical sequence.
    private bool TryWriteScalarBidiGlyphs(IReadOnlyList<PdfGlyphInfo> glyphs, string text,
        PdfTextShowCommand command, double fontSize, double textRise) {
        // Multi-character glyph mappings (including opt-in Latin ligatures) do not
        // have scalar fragment geometry; keep the exact logical replacement scope.
        foreach (PdfGlyphInfo glyph in glyphs)
            if (glyph.UnicodeText.Length != 1 && !(glyph.UnicodeText.Length == 2 && char.IsSurrogatePair(glyph.UnicodeText, 0))) return false;
        int[] scalars = OfficeUnicodeBidi.ToScalars(text, out _);
        foreach (int scalar in scalars)
            if (OfficeUnicodeBidi.GetClass(scalar) is not (OfficeBidiClass.L or OfficeBidiClass.R or OfficeBidiClass.AL)) return false;
        int[] levels = PdfUnicodeScalarTextShaper.ResolveGlyphLevels(text, glyphs);
        foreach (int level in levels) if (level > 1) return false;
        for (int start = 0; start < glyphs.Count;) {
            int end = start + 1;
            while (end < glyphs.Count && levels[end] == levels[start]) end++;
            _showEnd = -1;
            WriteLogicalGlyphs(glyphs, command, fontSize, textRise, start, end);
            _showEnd = -1;
            start = end;
        }
        return true;
    }
}
