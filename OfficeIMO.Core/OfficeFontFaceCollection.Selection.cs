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
            var available = new List<(OfficeFontFace Face, int RegistrationIndex)>();
            for (int index = _faces.Count - 1; index >= 0; index--) {
                OfficeFontFace face = _faces[index];
                if (!MatchesFamily(face, family)) continue;
                available.Add((face, index));
            }
            available.Sort((left, right) => {
                int rank = OfficeFontFaceMatcher.Compare(left.Face.Descriptor, right.Face.Descriptor, descriptor);
                return rank != 0 ? rank : right.RegistrationIndex.CompareTo(left.RegistrationIndex);
            });
            foreach ((OfficeFontFace face, _) in available) {
                if (added.Add(face)) result.Add(face);
            }
        }
        return result;
    }

}
