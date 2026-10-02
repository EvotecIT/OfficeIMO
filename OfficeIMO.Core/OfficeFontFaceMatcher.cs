using System;

namespace OfficeIMO.Drawing;

/// <summary>Shared CSS face ordering for registered and installed font candidates.</summary>
internal static class OfficeFontFaceMatcher {
    internal static int Compare(
        OfficeFontFaceDescriptor left,
        OfficeFontFaceDescriptor right,
        OfficeFontFaceDescriptor requested) {
        int comparison = CompareStretch(left.StretchPercent, right.StretchPercent, requested.StretchPercent);
        if (comparison != 0) return comparison;
        comparison = CompareSlant(left, right, requested);
        if (comparison != 0) return comparison;
        return CompareWeight(left.Weight, right.Weight, requested.Weight);
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
