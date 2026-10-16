using System;
using System.Collections.Generic;

namespace OfficeIMO.Drawing;

internal static class OfficeFontFamilyParser {
    internal const int DefaultMaximumCandidates = 32;
    internal const int DefaultMaximumFamilyNameLength = 256;

    /// <summary>Serializes one decoded family as a CSS string without turning literal commas into fallback separators.</summary>
    internal static string Quote(string familyName) =>
        "\"" + familyName.Replace("\\", "\\\\").Replace("\"", "\\\"") + "\"";

    internal static List<string> Parse(
        string? familyNames,
        int maximumCandidates = DefaultMaximumCandidates,
        int maximumFamilyNameLength = DefaultMaximumFamilyNameLength,
        bool allowSemicolonSeparators = false) {
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
                } else if (current == ',' || (allowSemicolonSeparators && current == ';')) {
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
                index++;
                current = value[index];
                if (TryGetHexDigit(current, out int digit)) {
                    int scalar = digit;
                    int digits = 1;
                    while (digits < 6 && index + 1 < end && TryGetHexDigit(value[index + 1], out digit)) {
                        scalar = scalar * 16 + digit;
                        digits++;
                        index++;
                    }
                    if (index + 1 < end && IsCssWhitespace(value[index + 1])) {
                        index++;
                        if (value[index] == '\r' && index + 1 < end && value[index + 1] == '\n') index++;
                    }
                    if (scalar <= 0 || scalar > 0x10FFFF || scalar is >= 0xD800 and <= 0xDFFF) scalar = 0xFFFD;
                    if (scalar > char.MaxValue) {
                        // Keep the decoded-name ceiling in UTF-16 units and never retain half a pair.
                        if (written + 2 > cleaned.Length) break;
                        scalar -= 0x10000;
                        cleaned[written++] = (char)(0xD800 + (scalar >> 10));
                        cleaned[written++] = (char)(0xDC00 + (scalar & 0x3FF));
                        continue;
                    }
                    current = (char)scalar;
                } else if (current is '\r' or '\n' or '\f') {
                    if (current == '\r' && index + 1 < end && value[index + 1] == '\n') index++;
                    continue;
                }
            }
            cleaned[written++] = current;
        }
        return new string(cleaned, 0, written);
    }

    private static bool TryGetHexDigit(char value, out int digit) {
        digit = value is >= '0' and <= '9' ? value - '0'
            : value is >= 'a' and <= 'f' ? value - 'a' + 10
            : value is >= 'A' and <= 'F' ? value - 'A' + 10 : -1;
        return digit >= 0;
    }

    private static bool IsCssWhitespace(char value) => value is ' ' or '\t' or '\r' or '\n' or '\f';

    private static void TrimBounds(string value, ref int start, ref int end) {
        while (start < end && char.IsWhiteSpace(value[start])) start++;
        while (end > start && char.IsWhiteSpace(value[end - 1])) end--;
    }
}
