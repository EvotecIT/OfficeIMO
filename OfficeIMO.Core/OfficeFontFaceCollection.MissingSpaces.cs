using System;
using System.Collections.Generic;

namespace OfficeIMO.Drawing;

public sealed partial class OfficeFontFaceCollection {
    private static bool CanInheritWhitespace(
        string element,
        string? currentFamily,
        IReadOnlyList<OfficeFontFace> candidates,
        HashSet<string> explicitlySelectedFaces) {
        if (currentFamily == null || !IsWhitespace(element)) return false;
        if (!OfficeMissingSpaceMetrics.ContainsEmSpace(element)) return true;
        foreach (OfficeFontFace candidate in candidates) {
            if (!string.Equals(candidate.ResourceFamilyName, currentFamily, StringComparison.OrdinalIgnoreCase)) continue;
            bool explicitlySelected = explicitlySelectedFaces.Contains(candidate.ResourceFamilyName);
            if (explicitlySelected ? candidate.HasGlyphs(element) : candidate.Covers(element)) return true;
        }
        return false;
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
