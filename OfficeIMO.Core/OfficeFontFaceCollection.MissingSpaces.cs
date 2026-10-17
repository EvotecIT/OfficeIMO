using System;
using System.Collections.Generic;

namespace OfficeIMO.Drawing;

public sealed partial class OfficeFontFaceCollection {
    /// <summary>
    /// Selects a complete missing-EM grapheme only after mapped-glyph candidates
    /// are exhausted. Generated run resources retain the original range constraint;
    /// only a caller's explicit resource request bypasses it.
    /// </summary>
    private static OfficeFontFace? ResolveSynthesizedSpaceFace(
        string element,
        string? currentFamily,
        IReadOnlyList<OfficeFontFace> candidates,
        HashSet<string> explicitlySelectedFaces) {
        OfficeFontFace? firstEligible = null;
        foreach (OfficeFontFace candidate in candidates) {
            bool explicitlySelected = explicitlySelectedFaces.Contains(candidate.ResourceFamilyName);
            if (!CoversPlannedText(candidate, element, requireUnicodeRange: !explicitlySelected, allowSpaceSynthesis: true)) continue;
            if (string.Equals(candidate.ResourceFamilyName, currentFamily, StringComparison.OrdinalIgnoreCase)) return candidate;
            firstEligible ??= candidate;
        }
        return firstEligible;
    }

    private static bool CanInheritWhitespace(
        string element,
        OfficeFontFace? currentFace,
        HashSet<string> explicitlySelectedFaces) {
        if (currentFace == null || !IsWhitespace(element)) return false;
        if (!OfficeMissingSpaceMetrics.ContainsEmSpace(element)) return true;
        bool explicitlySelected = explicitlySelectedFaces.Contains(currentFace.ResourceFamilyName);
        return explicitlySelected ? currentFace.HasGlyphs(element) : currentFace.Covers(element);
    }

    private static bool IsWhitespace(string value) {
        for (int index = 0; index < value.Length; index++) {
            if (!char.IsWhiteSpace(value[index])) return false;
        }
        return value.Length > 0;
    }

    private static bool CoversPlannedText(
        OfficeFontFace face,
        string text,
        bool requireUnicodeRange,
        bool allowSpaceSynthesis = false) {
        foreach (string element in OfficeTextElements.Enumerate(text)) {
            if (IsWhitespace(element) && !OfficeMissingSpaceMetrics.ContainsEmSpace(element)) continue;
            if (requireUnicodeRange ? face.Covers(element) : face.HasGlyphs(element)) continue;
            if (!allowSpaceSynthesis || requireUnicodeRange && !face.UnicodeRanges.ContainsFontCoverageText(element) ||
                !OfficeMissingSpaceMetrics.CoversWithSynthesis(face.Program, element)) return false;
        }
        return text.Length > 0;
    }
}
