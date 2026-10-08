namespace OfficeIMO.Pdf;

public sealed partial class PdfOptions {
    private HashSet<PdfStandardFont>? _fallbackOwnedFontFamilies;
    private HashSet<string>? _automaticFallbackOwnedNamedFamilyKeys;

    internal bool HasCallerOwnedEmbeddedFontFamily(PdfStandardFont family) {
        PdfStandardFont normalized = PdfStandardFontMapper.GetFontFamily(family);
        return _fallbackOwnedFontFamilies?.Contains(normalized) != true &&
            _embeddedFonts?.Keys.Any(font => PdfStandardFontMapper.GetFontFamily(font) == normalized) == true;
    }

    private void MergeResolvedFallbackFontMappingsFrom(PdfOptions nested) {
        if (nested._usedEmbeddedFallbackFontSlots == null || nested._fallbackOwnedFontFamilies == null) return;
        foreach (PdfStandardFont family in nested._usedEmbeddedFallbackFontSlots.OrderBy(value => value)) {
            if (!nested._fallbackOwnedFontFamilies.Contains(family)) continue;
            if (HasCallerOwnedEmbeddedFontFamily(family))
                throw new InvalidOperationException("Nested fallback fonts cannot replace a caller-owned font family.");

            for (int style = 0; style < 4; style++) {
                PdfStandardFont font = PdfStandardFontMapper.GetStyledFont(family,
                    bold: (style & 1) != 0, italic: (style & 2) != 0);
                if (!nested.TryGetEmbeddedStandardFont(font, out PdfEmbeddedFont? embedded) || embedded == null) continue;
                if (IsEmbeddedFallbackFontFamilySlotUsed(family)
                    && TryGetEmbeddedStandardFont(font, out PdfEmbeddedFont? existing) && existing != null
                    && (!EmbeddedFontMatches(existing, embedded.DataSnapshot, embedded.FontName)
                        || existing.SyntheticOblique != embedded.SyntheticOblique))
                    throw new InvalidOperationException("Nested fallback fonts cannot replace a font family already used by the page.");

                // The page resource owner needs the mapping, not only the child
                // program's glyph usage. Immutable snapshots can be shared safely.
                EmbedStandardFontSnapshot(font, embedded.DataSnapshot, embedded.FontName, embedded.SyntheticOblique);
            }
            (_fallbackOwnedFontFamilies ??= new HashSet<PdfStandardFont>()).Add(family);
            MarkEmbeddedFallbackFontFamilySlotUsed(family);
        }
    }

    private void ReleasePreviousFallbackFontSlots(PdfEmbeddedFontFallbackSet? previous) {
        if (_automaticFallbackOwnedNamedFamilyKeys != null) {
            foreach (string key in _automaticFallbackOwnedNamedFamilyKeys) {
                if (_namedFontFamilies?.Remove(key) == true) {
                    RemoveNamedFontProgramCache(key);
                }
            }
            _automaticFallbackOwnedNamedFamilyKeys.Clear();
        }
        if (previous == null || previous.UsesNamedFontFamilies || _fallbackOwnedFontFamilies == null) return;
        // Requested slots may change. Reclaim every mapping still owned by the old
        // set, including the writer's resolved slots, while retaining caller faces.
        foreach (PdfStandardFont family in _fallbackOwnedFontFamilies.ToArray()) {
            if (!previous.Candidates.Any(candidate => IsRegisteredFallbackFontFamily(family, candidate))) continue;
            for (int style = 0; style < 4; style++) {
                PdfStandardFont font = PdfStandardFontMapper.GetStyledFont(family,
                    bold: (style & 1) != 0, italic: (style & 2) != 0);
                _embeddedFonts?.Remove(font);
                _embeddedFontPrograms?.Remove(font);
                _embeddedOpenTypeCffFontPrograms?.Remove(font);
                _embeddedFontProgramFailures?.Remove(font);
                ClearReportedEmbeddedFontProgramFailure(font);
            }
            _fallbackOwnedFontFamilies.Remove(family);
            _usedEmbeddedFallbackFontSlots?.Remove(family);
        }
    }
}
