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
        int maximumSourceCharacters = checked(maximumCandidates * (maximumFamilyNameLength * 2 + 3));
        int scanEnd = Math.Min(familyNames!.Length, maximumSourceCharacters);
        int segmentStart = 0;
        while (segmentStart < scanEnd && families.Count < maximumCandidates) {
            int segmentEnd = segmentStart;
            char quote = '\0';
            bool escaped = false;
            while (segmentEnd < scanEnd) {
                char current = familyNames[segmentEnd];
                if (escaped) {
                    escaped = false;
                } else if (current == '\\') {
                    escaped = true;
                } else if (quote != '\0') {
                    if (current == quote) quote = '\0';
                } else if (current == '\'' || current == '"') {
                    quote = current;
                } else if (current == ',') {
                    break;
                }
                segmentEnd++;
            }
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

        int length = end - start;
        if (length <= 0) return string.Empty;
        int slash = value.IndexOf('\\', start, length);
        if (slash < 0) return value.Substring(start, Math.Min(length, maximumLength));

        var cleaned = new char[Math.Min(length, maximumLength)];
        int written = 0;
        for (int index = start; index < end && written < cleaned.Length; index++) {
            char current = value[index];
            if (current == '\\' && index + 1 < end) {
                char escaped = value[index + 1];
                if (escaped == ',' || escaped == '\\' || escaped == '\'' || escaped == '"') {
                    current = escaped;
                    index++;
                }
            }
            cleaned[written++] = current;
        }
        return new string(cleaned, 0, written);
    }

    private static void TrimBounds(string value, ref int start, ref int end) {
        while (start < end && char.IsWhiteSpace(value[start])) start++;
        while (end > start && char.IsWhiteSpace(value[end - 1])) end--;
    }
}
