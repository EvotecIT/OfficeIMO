using System;
using System.Collections.Generic;

namespace OfficeIMO.Drawing;

public sealed partial class OfficeFontFaceCollection {
    private IReadOnlyList<OfficeFontFace> ResolveFallbackCandidates(string familyNames, OfficeFontStyle style) =>
        ResolveFallbackCandidates(familyNames, OfficeFontFaceDescriptor.FromStyle(style));

    private IReadOnlyList<OfficeFontFace> ResolveFallbackCandidates(
        string familyNames,
        OfficeFontFaceDescriptor descriptor) {
        if (_faces.Count == 0) return Array.Empty<OfficeFontFace>();

        var result = new List<OfficeFontFace>();
        var added = new HashSet<OfficeFontFace>();
        var families = new List<string>();
        var addedFamilies = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (string family in OfficeFontFamilyParser.Parse(familyNames)) {
            if (addedFamilies.Add(family)) families.Add(family);
        }
        foreach (string family in _fallbackFamilies) {
            if (addedFamilies.Add(family)) families.Add(family);
        }
        foreach (string family in families) {
            List<(OfficeFontFace Face, int RegistrationIndex)> available = ResolveFamilyCandidates(family, descriptor);
            foreach ((OfficeFontFace face, _) in available) {
                if (added.Add(face)) result.Add(face);
            }
        }
        return result;
    }

    /// <summary>Resolves a complete text element within one requested family, without entering later fallback families.</summary>
    internal OfficeFontFace? ResolveFaceInFamily(string text, string family, OfficeFontFaceDescriptor descriptor) {
        foreach ((OfficeFontFace face, _) in ResolveFamilyCandidates(family, descriptor)) {
            bool explicitResource = string.Equals(face.ResourceFamilyName, family, StringComparison.OrdinalIgnoreCase);
            if (explicitResource ? face.HasGlyphs(text) : face.Covers(text)) return face;
        }
        return null;
    }

    private List<(OfficeFontFace Face, int RegistrationIndex)> ResolveFamilyCandidates(string family, OfficeFontFaceDescriptor descriptor) {
        var available = new List<(OfficeFontFace Face, int RegistrationIndex)>();
        for (int index = _faces.Count - 1; index >= 0; index--) {
            OfficeFontFace face = _faces[index];
            if (!MatchesFamily(face, family)) continue;
            available.Add((face, index));
        }
        available.Sort((left, right) => {
            if (OfficeSystemFontFamilyAliases.IsMath(family)) {
                // An explicitly supplied generic face wins; otherwise keep the canonical
                // mathematical family order before style matching within a family.
                int familyRank(OfficeFontFace face) {
                    if (string.Equals(face.FamilyName, family, StringComparison.OrdinalIgnoreCase)
                        || string.Equals(face.ResourceFamilyName, family, StringComparison.OrdinalIgnoreCase)) return 0;
                    return OfficeSystemFontFamilyAliases.MathFamilyRank(face.FamilyName);
                }
                int familyComparison = familyRank(left.Face).CompareTo(familyRank(right.Face));
                if (familyComparison != 0) return familyComparison;
            }
            int rank = CompareFaceSelection(left.Face, right.Face, descriptor);
            return rank != 0 ? rank : right.RegistrationIndex.CompareTo(left.RegistrationIndex);
        });
        return available;
    }

    private static int CompareFaceSelection(
        OfficeFontFace left,
        OfficeFontFace right,
        OfficeFontFaceDescriptor requested) => OfficeFontFaceMatcher.Compare(left.Descriptor, right.Descriptor, requested);

    /// <summary>Ranks descriptors without requiring decoded programs, so consumers can attribute unavailable preferred faces.</summary>
    internal static int CompareFaceDescriptors(
        OfficeFontFaceDescriptor left,
        OfficeFontFaceDescriptor right,
        OfficeFontFaceDescriptor requested) => OfficeFontFaceMatcher.Compare(left, right, requested);
}
