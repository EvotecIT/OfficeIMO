using System;
using System.Collections.Generic;

namespace OfficeIMO.Drawing;

/// <summary>
/// Supplies a blank one-em advance for absent EM SPACE and EM QUAD glyphs.
/// Actual glyph metrics and provider-owned programs remain authoritative.
/// </summary>
internal static class OfficeMissingSpaceMetrics {
    internal static bool IsEmSpace(int scalar) => scalar == 0x2001 || scalar == 0x2003;

    internal static bool ContainsEmSpace(string text) =>
        text.IndexOf('\u2001') >= 0 || text.IndexOf('\u2003') >= 0;

    internal static bool ContainsUnmappedEmSpace(IReadOnlyList<int> glyphs, IReadOnlyList<int> scalars) {
        for (int index = 0; index < glyphs.Count; index++) {
            if (glyphs[index] == 0 && IsEmSpace(scalars[index])) return true;
        }
        return false;
    }

    // Called only after the cheap missing-scalar check, so ordinary text creates
    // neither an extra array nor a font-mapping delegate.
    internal static bool[]? FindMissingEmSpaces(
        IReadOnlyList<int> glyphs,
        IReadOnlyList<int> scalars,
        Func<int, bool> hasMappedGlyph) {
        bool[]? missing = null;
        for (int index = 0; index < glyphs.Count; index++) {
            if (glyphs[index] != 0 || !IsEmSpace(scalars[index]) || hasMappedGlyph(scalars[index])) continue;
            missing ??= new bool[glyphs.Count];
            missing[index] = true;
        }
        return missing;
    }

    internal static bool IsSynthesized(bool[]? missing, int index) => missing != null && missing[index];

    internal static bool CoversWithSynthesis(IOfficeFontProgram program, string text) {
        if (!(program is OfficeTrueTypeFont) && !(program is OfficeOpenTypeCffFont) || !ContainsEmSpace(text)) return false;
        // Keep the complete grapheme as one fallback unit. Every other scalar,
        // including attached marks, must still be covered by the selected face.
        string remaining = text;
        if (!program.HasGlyphs("\u2001")) remaining = remaining.Replace("\u2001", string.Empty);
        if (!program.HasGlyphs("\u2003")) remaining = remaining.Replace("\u2003", string.Empty);
        return program.HasGlyphs(remaining);
    }
}
