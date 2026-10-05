using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal sealed partial class PdfTrueTypeFontProgram {
    private const int FirstAsciiCharacter = 32;
    private const int AsciiCharacterCount = 95;
    private ulong _usedAsciiLow, _usedAsciiHigh;

    // Separate CIDs give ordinary characters stable Unicode mappings even when
    // a later ligature or Unicode alias changes the mapping of their shared GID.
    // Original GID-based CIDs remain available for every other shaping route.
    internal bool TryCreateAsciiTextShowCommand(PdfGlyphRun run, out PdfTextShowCommand command) {
        command = null!;
        if (!run.PreserveGlyphUnicode || run.ActualText != null || run.HasPositioning ||
            run.HasMissingGlyphs || run.Glyphs.Count == 0 || GlyphCount > 65536 - AsciiCharacterCount ||
            (run.Direction != OfficeTextDirection.Auto && run.Direction != OfficeTextDirection.LeftToRight)) {
            return false;
        }

        ulong low = 0, high = 0;
        for (int index = 0; index < run.Glyphs.Count; index++) {
            PdfGlyphInfo glyph = run.Glyphs[index];
            if (glyph.UnicodeText.Length != 1 || glyph.TextIndex != index || glyph.LogicalClusterStart != index) return false;
            int scalar = glyph.UnicodeText[0];
            int offset = scalar - FirstAsciiCharacter;
            if (offset < 0 || offset >= AsciiCharacterCount ||
                !_cmap.TryGetValue(scalar, out int glyphId) || glyphId <= 0 || glyphId >= GlyphCount || glyphId != glyph.GlyphId) return false;
            if (offset < 64) low |= 1UL << offset;
            else high |= 1UL << (offset - 64);
        }

        var hex = PdfGlyphRun.RentHexBuilder(run.Glyphs.Count * 4);
        for (int index = 0; index < run.Glyphs.Count; index++)
            PdfGlyphRun.AppendGlyphHex(hex, GetAsciiCharacterCode(run.Glyphs[index].UnicodeText[0]));
        string glyphHex = PdfGlyphRun.ReturnHexBuilder(hex);
        lock (_usageLock) {
            _usedAsciiLow |= low;
            _usedAsciiHigh |= high;
        }
        command = new PdfTextShowCommand(glyphHex, advanceWidth1000: run.TotalAdvanceWidth1000, visualGlyphs: run.Glyphs);
        return true;
    }

    private int GetAsciiCharacterCode(int scalar) => GlyphCount + scalar - FirstAsciiCharacter;

    internal IReadOnlyList<(int CharacterCode, int GlyphId, string UnicodeText)> GetUsedAsciiCharacterMappings() {
        var result = new List<(int, int, string)>();
        lock (_usageLock) {
            for (int offset = 0; offset < AsciiCharacterCount; offset++) {
                if (!IsAsciiCharacterUsed(offset)) continue;
                int scalar = offset + FirstAsciiCharacter;
                result.Add((GetAsciiCharacterCode(scalar), _cmap[scalar], ((char)scalar).ToString()));
            }
        }
        return result;
    }

    internal IReadOnlyList<(int CharacterCode, string UnicodeText)> GetCharacterCodeToUnicodeMappings() {
        var result = new List<(int, string)>(GetGlyphToUnicodeMappings());
        foreach (var mapping in GetUsedAsciiCharacterMappings())
            result.Add((mapping.CharacterCode, mapping.UnicodeText));
        return result;
    }

    // Snapshot the source before locking the destination: option-context merges
    // must not acquire two usage locks in opposing order.
    internal void MergeAsciiCharacterUsageFrom(PdfTrueTypeFontProgram source) {
        if (ReferenceEquals(this, source)) return;
        if (!_subsetFontFingerprint.Equals(source._subsetFontFingerprint))
            throw new InvalidOperationException("Character-code usage can only be merged for the same TrueType font program.");
        ulong low, high;
        lock (source._usageLock) { low = source._usedAsciiLow; high = source._usedAsciiHigh; }
        lock (_usageLock) { _usedAsciiLow |= low; _usedAsciiHigh |= high; }
    }

    internal byte[]? BuildCidToGlyphMap() {
        lock (_usageLock) {
            if ((_usedAsciiLow | _usedAsciiHigh) == 0) return null;
            // Unused entries stay zero, keeping large-font maps sparse and
            // compressible. Only painted original CIDs need identity entries.
            var map = new byte[(GlyphCount + AsciiCharacterCount) * 2];
            foreach (int glyphId in _usedGlyphIds) {
                if (glyphId >= 0 && glyphId < GlyphCount) WriteCidGlyph(map, glyphId, glyphId);
            }
            for (int offset = 0; offset < AsciiCharacterCount; offset++) {
                if (!IsAsciiCharacterUsed(offset)) continue;
                int scalar = offset + FirstAsciiCharacter;
                WriteCidGlyph(map, GetAsciiCharacterCode(scalar), _cmap[scalar]);
            }
            return map;
        }
    }

    private bool IsAsciiCharacterUsed(int offset) => offset < 64
        ? (_usedAsciiLow & (1UL << offset)) != 0
        : (_usedAsciiHigh & (1UL << (offset - 64))) != 0;

    private static void WriteCidGlyph(byte[] map, int characterCode, int glyphId) {
        map[characterCode * 2] = (byte)(glyphId >> 8);
        map[characterCode * 2 + 1] = (byte)glyphId;
    }
}
