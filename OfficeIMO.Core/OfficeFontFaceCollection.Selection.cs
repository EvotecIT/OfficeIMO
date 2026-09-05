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
            var available = new List<OfficeFontFace>();
            for (int index = _faces.Count - 1; index >= 0; index--) {
                OfficeFontFace face = _faces[index];
                if (!MatchesFamily(face, family)) continue;
                int insertionIndex = available.Count;
                for (int candidateIndex = 0; candidateIndex < available.Count; candidateIndex++) {
                    if (CompareFaceSelection(face, available[candidateIndex], descriptor) < 0) {
                        insertionIndex = candidateIndex;
                        break;
                    }
                }
                available.Insert(insertionIndex, face);
            }
            foreach (OfficeFontFace face in available) {
                if (added.Add(face)) result.Add(face);
            }
        }
        return result;
    }

    private static int CompareFaceSelection(
        OfficeFontFace left,
        OfficeFontFace right,
        OfficeFontFaceDescriptor requested) {
        int comparison = CompareStretch(left.Descriptor.StretchPercent, right.Descriptor.StretchPercent, requested.StretchPercent);
        if (comparison != 0) return comparison;
        comparison = CompareSlant(left.Descriptor, right.Descriptor, requested);
        if (comparison != 0) return comparison;
        return CompareWeight(left.Descriptor.Weight, right.Descriptor.Weight, requested.Weight);
    }

    private static int CompareStretch(double left, double right, double requested) {
        (int leftZone, double leftDistance) = StretchRank(left, requested);
        (int rightZone, double rightDistance) = StretchRank(right, requested);
        int comparison = leftZone.CompareTo(rightZone);
        return comparison != 0 ? comparison : leftDistance.CompareTo(rightDistance);
    }

    private static (int Zone, double Distance) StretchRank(double candidate, double requested) {
        bool preferredDirection = requested <= 100D ? candidate <= requested : candidate >= requested;
        return (preferredDirection ? 0 : 1, Math.Abs(candidate - requested));
    }

    private static int CompareSlant(
        OfficeFontFaceDescriptor left,
        OfficeFontFaceDescriptor right,
        OfficeFontFaceDescriptor requested) {
        (int leftZone, double leftDistance) = SlantRank(left, requested);
        (int rightZone, double rightDistance) = SlantRank(right, requested);
        int comparison = leftZone.CompareTo(rightZone);
        return comparison != 0 ? comparison : leftDistance.CompareTo(rightDistance);
    }

    private static (int Zone, double Distance) SlantRank(
        OfficeFontFaceDescriptor candidate,
        OfficeFontFaceDescriptor requested) {
        if (candidate.Slant == requested.Slant) {
            return (0, requested.Slant == OfficeFontSlant.Oblique
                ? Math.Abs(candidate.ObliqueAngleDegrees - requested.ObliqueAngleDegrees)
                : 0D);
        }
        if (requested.Slant == OfficeFontSlant.Normal) {
            return (candidate.Slant == OfficeFontSlant.Oblique ? 1 : 2, 0D);
        }
        return (candidate.Slant == OfficeFontSlant.Normal ? 2 : 1, 0D);
    }

    private static int CompareWeight(int left, int right, int requested) {
        (int leftZone, int leftDistance) = WeightRank(left, requested);
        (int rightZone, int rightDistance) = WeightRank(right, requested);
        int comparison = leftZone.CompareTo(rightZone);
        return comparison != 0 ? comparison : leftDistance.CompareTo(rightDistance);
    }

    private static (int Zone, int Distance) WeightRank(int candidate, int requested) {
        if (requested >= 400 && requested <= 500) {
            if (candidate >= requested && candidate <= 500) return (0, candidate - requested);
            if (candidate < requested) return (1, requested - candidate);
            return (2, candidate - 500);
        }
        if (requested < 400) {
            return candidate <= requested
                ? (0, requested - candidate)
                : (1, candidate - requested);
        }
        return candidate >= requested
            ? (0, candidate - requested)
            : (1, requested - candidate);
    }
}
