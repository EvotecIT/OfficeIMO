namespace OfficeIMO.Pdf;

public sealed partial class PdfOptions {
    internal void MergeFontProgramUsageFrom(PdfOptions nested) {
        if (ReferenceEquals(this, nested)) return;
        if (nested._embeddedFontPrograms != null) {
            foreach (var entry in nested._embeddedFontPrograms) {
                if (TryGetEmbeddedStandardFontProgramForGeneration(entry.Key, out _, out PdfTrueTypeFontProgram? target) && target != null) {
                    foreach ((int glyphId, string unicodeText) in entry.Value.GetGlyphToUnicodeMappings())
                        target.RecordGlyphUsage(glyphId, unicodeText);
                }
            }
        }
        if (nested._embeddedOpenTypeCffFontPrograms != null) {
            foreach (var entry in nested._embeddedOpenTypeCffFontPrograms) {
                if (TryGetEmbeddedStandardOpenTypeCffFontProgramForGeneration(entry.Key, out _, out PdfOpenTypeCffFontProgram? target) && target != null) {
                    foreach ((int glyphId, string unicodeText) in entry.Value.GetGlyphToUnicodeMappings())
                        target.RecordGlyphUsage(glyphId, unicodeText);
                }
            }
        }
        if (nested._namedFontPrograms != null) {
            foreach (var entry in nested._namedFontPrograms) {
                if (TryGetNamedFontProgramForGeneration(entry.Key, out PdfTrueTypeFontProgram? target) && target != null) {
                    foreach ((int glyphId, string unicodeText) in entry.Value.GetGlyphToUnicodeMappings())
                        target.RecordGlyphUsage(glyphId, unicodeText);
                }
            }
        }
        if (nested._namedOpenTypeCffFontPrograms != null) {
            foreach (var entry in nested._namedOpenTypeCffFontPrograms) {
                if (TryGetNamedOpenTypeCffFontProgramForGeneration(entry.Key, out PdfOpenTypeCffFontProgram? target) && target != null) {
                    foreach ((int glyphId, string unicodeText) in entry.Value.GetGlyphToUnicodeMappings())
                        target.RecordGlyphUsage(glyphId, unicodeText);
                }
            }
        }
    }
}
