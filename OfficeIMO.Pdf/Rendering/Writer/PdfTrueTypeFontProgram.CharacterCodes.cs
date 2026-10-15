namespace OfficeIMO.Pdf;

internal sealed partial class PdfTrueTypeFontProgram {
    // CIDs are glyph ids. Pins keep each emitted glyph's ToUnicode entry deterministic, and
    // alias codes (glyph count + scalar) keep characters that share a nominal glyph distinct.
    internal IPdfGlyphTextPinner TextPinner => _textPins;

    internal IReadOnlyList<(int CharacterCode, int GlyphId, string UnicodeText)> GetAliasCharacterMappings() => _textPins.GetAliases();

    internal IReadOnlyList<(int CharacterCode, string UnicodeText)> GetCharacterCodeToUnicodeMappings() {
        var result = new List<(int, string)>(GetGlyphToUnicodeMappings());
        foreach (var alias in _textPins.GetAliases()) result.Add((alias.Code, alias.UnicodeText));
        return result;
    }

    internal void MergeTextPinsFrom(PdfTrueTypeFontProgram source) {
        if (ReferenceEquals(this, source)) return;
        if (!_subsetFontFingerprint.Equals(source._subsetFontFingerprint))
            throw new InvalidOperationException("Glyph text pins can only be merged for the same TrueType font program.");
        _textPins.MergeFrom(source._textPins);
    }

    /// <summary>BaseFont and FontName of the embedded program: subset-tagged once glyphs are renumbered.</summary>
    internal string SubsetFontName {
        get {
            int[]? glyphOrder = GetCompactSubsetGlyphOrder();
            return glyphOrder == null ? FontName : PdfSubsetTag.Apply(FontName, glyphOrder);
        }
    }

    // Maps painted CIDs (glyph ids and alias codes) into the renumbered subset program.
    internal byte[]? BuildCidToGlyphMap() {
        IReadOnlyList<(int Code, int GlyphId, string UnicodeText)> aliases = _textPins.GetAliases();
        int[]? glyphOrder = GetCompactSubsetGlyphOrder();
        if (aliases.Count == 0 && glyphOrder == null) return null;
        var subsetGlyphIds = new Dictionary<int, int>();
        if (glyphOrder != null) {
            for (int index = 0; index < glyphOrder.Length; index++) subsetGlyphIds[glyphOrder[index]] = index;
        }
        int Subset(int glyphId) => glyphOrder == null ? glyphId : subsetGlyphIds.TryGetValue(glyphId, out int mapped) ? mapped : 0;

        IReadOnlyList<int> usedGlyphIds = GetUsedGlyphIds();
        int maximumCode = 0;
        bool identity = aliases.Count == 0;
        foreach (int glyphId in usedGlyphIds) {
            maximumCode = Math.Max(maximumCode, glyphId);
            identity &= Subset(glyphId) == glyphId;
        }
        if (identity) return null;
        foreach (var alias in aliases) maximumCode = Math.Max(maximumCode, alias.Code);
        // Unused entries stay zero, keeping large-font maps sparse and compressible.
        var map = new byte[(maximumCode + 1) * 2];
        foreach (int glyphId in usedGlyphIds) {
            if (glyphId >= 0 && glyphId < GlyphCount) WriteCidGlyph(map, glyphId, Subset(glyphId));
        }
        foreach (var alias in aliases) WriteCidGlyph(map, alias.Code, Subset(alias.GlyphId));
        return map;
    }

    private static void WriteCidGlyph(byte[] map, int characterCode, int glyphId) {
        map[characterCode * 2] = (byte)(glyphId >> 8);
        map[characterCode * 2 + 1] = (byte)glyphId;
    }
}
