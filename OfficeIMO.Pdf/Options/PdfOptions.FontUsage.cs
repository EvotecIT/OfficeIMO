namespace OfficeIMO.Pdf;

public sealed partial class PdfCreateOptions {
    private readonly object _fallbackFontUsageLock = new();
    private HashSet<PdfStandardFont>? _usedEmbeddedFallbackFontSlots;

    /// <summary>Records font usage when immutable document content is written concurrently.</summary>
    internal void MarkEmbeddedFallbackFontFamilySlotUsed(PdfStandardFont font) {
        PdfStandardFont family = PdfStandardFontMapper.GetFontFamily(font);
        lock (_fallbackFontUsageLock) {
            (_usedEmbeddedFallbackFontSlots ??= new HashSet<PdfStandardFont>()).Add(family);
        }
    }

    internal bool IsEmbeddedFallbackFontFamilySlotUsed(PdfStandardFont font) {
        PdfStandardFont family = PdfStandardFontMapper.GetFontFamily(font);
        lock (_fallbackFontUsageLock) {
            return _usedEmbeddedFallbackFontSlots?.Contains(family) == true;
        }
    }

    /// <summary>Copies usage under the writer lock; callers enumerate their own snapshot.</summary>
    private HashSet<PdfStandardFont>? CloneUsedEmbeddedFallbackFontSlots() {
        lock (_fallbackFontUsageLock) {
            return _usedEmbeddedFallbackFontSlots == null
                ? null : new HashSet<PdfStandardFont>(_usedEmbeddedFallbackFontSlots);
        }
    }

    private void RemoveUsedEmbeddedFallbackFontSlot(PdfStandardFont family) {
        lock (_fallbackFontUsageLock) {
            _usedEmbeddedFallbackFontSlots?.Remove(family);
        }
    }

    private void ClearUsedEmbeddedFallbackFontSlots() {
        lock (_fallbackFontUsageLock) {
            _usedEmbeddedFallbackFontSlots?.Clear();
        }
    }

    internal void MergeFontProgramUsageFrom(PdfCreateOptions nested) {
        if (ReferenceEquals(this, nested)) return;
        MergeResolvedFallbackFontMappingsFrom(nested);
        if (nested._embeddedFontPrograms != null) {
            foreach (var entry in nested._embeddedFontPrograms) {
                if (TryGetEmbeddedStandardFontProgramForGeneration(entry.Key, out _, out PdfTrueTypeFontProgram? target) && target != null) {
                    foreach ((int glyphId, string unicodeText) in entry.Value.GetUsedGlyphToUnicodeMappings())
                        target.RecordGlyphUsage(glyphId, unicodeText);
                    target.MergeTextPinsFrom(entry.Value);
                }
            }
        }
        if (nested._embeddedOpenTypeCffFontPrograms != null) {
            foreach (var entry in nested._embeddedOpenTypeCffFontPrograms) {
                if (TryGetEmbeddedStandardOpenTypeCffFontProgramForGeneration(entry.Key, out _, out PdfOpenTypeCffFontProgram? target) && target != null) {
                    foreach ((int glyphId, string unicodeText) in entry.Value.GetUsedGlyphToUnicodeMappings())
                        target.RecordGlyphUsage(glyphId, unicodeText);
                    target.MergeTextPinsFrom(entry.Value);
                }
            }
        }
        if (nested._namedFontPrograms != null) {
            foreach (var entry in nested._namedFontPrograms) {
                if (TryGetNamedFontProgramForGeneration(entry.Key, out PdfTrueTypeFontProgram? target) && target != null) {
                    foreach ((int glyphId, string unicodeText) in entry.Value.GetUsedGlyphToUnicodeMappings())
                        target.RecordGlyphUsage(glyphId, unicodeText);
                    target.MergeTextPinsFrom(entry.Value);
                }
            }
        }
        if (nested._namedOpenTypeCffFontPrograms != null) {
            foreach (var entry in nested._namedOpenTypeCffFontPrograms) {
                if (TryGetNamedOpenTypeCffFontProgramForGeneration(entry.Key, out PdfOpenTypeCffFontProgram? target) && target != null) {
                    foreach ((int glyphId, string unicodeText) in entry.Value.GetUsedGlyphToUnicodeMappings())
                        target.RecordGlyphUsage(glyphId, unicodeText);
                    target.MergeTextPinsFrom(entry.Value);
                }
            }
        }
    }
}
