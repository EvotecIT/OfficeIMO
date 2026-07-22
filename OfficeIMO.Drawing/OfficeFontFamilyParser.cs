using System;
using System.Collections.Generic;

namespace OfficeIMO.Drawing;

internal static class OfficeFontFamilyParser {
    internal const int DefaultMaximumCandidates = 32;
    internal const int DefaultMaximumFamilyNameLength = 256;

    internal static List<string> Parse(
        string? familyNames,
        int maximumCandidates = DefaultMaximumCandidates,
        int maximumFamilyNameLength = DefaultMaximumFamilyNameLength) {
        var families = new List<string>();
        if (string.IsNullOrEmpty(familyNames) || maximumCandidates < 1 || maximumFamilyNameLength < 1) {
            return families;
        }

        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        int maximumSourceCharacters = checked(maximumCandidates * (maximumFamilyNameLength + 1));
        int scanEnd = Math.Min(familyNames!.Length, maximumSourceCharacters);
        int segmentStart = 0;
        while (segmentStart < scanEnd && families.Count < maximumCandidates) {
            int segmentEnd = segmentStart;
            while (segmentEnd < scanEnd && familyNames[segmentEnd] != ',') segmentEnd++;
            string family = CleanSegment(familyNames, segmentStart, segmentEnd, maximumFamilyNameLength);
            if (family.Length > 0 && seen.Add(family)) families.Add(family);
            segmentStart = segmentEnd + 1;
        }

        return families;
    }

    private static string CleanSegment(string value, int start, int end, int maximumLength) {
        TrimBounds(value, ref start, ref end);
        while (end - start >= 2 &&
               ((value[start] == '"' && value[end - 1] == '"') ||
                (value[start] == '\'' && value[end - 1] == '\''))) {
            start++;
            end--;
            TrimBounds(value, ref start, ref end);
        }

        int length = Math.Min(end - start, maximumLength);
        return length > 0 ? value.Substring(start, length) : string.Empty;
    }

    private static void TrimBounds(string value, ref int start, ref int end) {
        while (start < end && char.IsWhiteSpace(value[start])) start++;
        while (end > start && char.IsWhiteSpace(value[end - 1])) end--;
    }
}
