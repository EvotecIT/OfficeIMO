namespace OfficeIMO.Pdf;

public sealed partial class PdfCreateOptions {
    private HashSet<PdfStandardFont>? _fallbackOwnedFontFamilies;
    private HashSet<string>? _automaticFallbackOwnedNamedFamilyKeys;

    internal bool HasCallerOwnedEmbeddedFontFamily(PdfStandardFont family) {
        PdfStandardFont normalized = PdfStandardFontMapper.GetFontFamily(family);
        ResolvePendingDefaultFontFamilies(normalized);
        return _fallbackOwnedFontFamilies?.Contains(normalized) != true &&
            HasEmbeddedFontKeyInFamily(normalized);
    }

    // Runs per measured and drawn text; a key loop avoids a closure and a boxed key enumerator per call.
    private bool HasEmbeddedFontKeyInFamily(PdfStandardFont normalized) {
        if (_embeddedFonts == null) return false;
        foreach (PdfStandardFont font in _embeddedFonts.Keys) {
            if (PdfStandardFontMapper.GetFontFamily(font) == normalized) return true;
        }
        return false;
    }

    internal bool HasEmbeddedFontInFamily(PdfStandardFont family) {
        PdfStandardFont normalized = PdfStandardFontMapper.GetFontFamily(family);
        ResolvePendingDefaultFontFamilies(normalized);
        return HasEmbeddedFontKeyInFamily(normalized);
    }

    private void MergeResolvedFallbackFontMappingsFrom(PdfCreateOptions nested) {
        MergeResolvedNamedFallbackFontMappingsFrom(nested);
        if (nested._fallbackOwnedFontFamilies == null) return;
        HashSet<PdfStandardFont>? usedSlots = nested.CloneUsedEmbeddedFallbackFontSlots();
        if (usedSlots == null) return;
        foreach (PdfStandardFont family in usedSlots.OrderBy(value => value)) {
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
                        || existing.SyntheticOblique != embedded.SyntheticOblique
                        || existing.SyntheticBold != embedded.SyntheticBold))
                    throw new InvalidOperationException("Nested fallback fonts cannot replace a font family already used by the page.");

                // The page resource owner needs the mapping, not only the child
                // program's glyph usage. Immutable snapshots can be shared safely.
                EmbedStandardFontSnapshot(font, embedded.DataSnapshot, embedded.FontName, embedded.SyntheticOblique, embedded.SyntheticBold);
            }
            (_fallbackOwnedFontFamilies ??= new HashSet<PdfStandardFont>()).Add(family);
            MarkEmbeddedFallbackFontFamilySlotUsed(family);
        }
    }

    private void ReleasePreviousFallbackFontSlots(PdfEmbeddedFontFallbackSet? previous) {
        if (_automaticFallbackOwnedNamedFamilyKeys != null) {
            foreach (string key in _automaticFallbackOwnedNamedFamilyKeys) {
                RemoveNamedFallbackPlans(key);
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
                ResolvePendingDefaultFontFamilies(font);
                _embeddedFonts?.Remove(font);
                _embeddedFontPrograms?.Remove(font);
                _embeddedOpenTypeCffFontPrograms?.Remove(font);
                _embeddedFontProgramFailures?.Remove(font);
                ClearReportedEmbeddedFontProgramFailure(font);
            }
            _fallbackOwnedFontFamilies.Remove(family);
            RemoveUsedEmbeddedFallbackFontSlot(family);
        }
    }
}
