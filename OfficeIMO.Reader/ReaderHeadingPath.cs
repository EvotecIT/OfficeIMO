using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeIMO.Reader;

/// <summary>Builds and parses escaped Reader heading paths.</summary>
public static class ReaderHeadingPath {
    private const string Separator = " > ";

    /// <summary>Combines heading titles while escaping literal backslashes and greater-than signs.</summary>
    public static string? Combine(IEnumerable<string?> headings) {
        if (headings == null) throw new ArgumentNullException(nameof(headings));
        string[] values = headings
            .Where(static heading => !string.IsNullOrWhiteSpace(heading))
            .Select(static heading => Escape(heading!.Trim()))
            .ToArray();
        return values.Length == 0 ? null : string.Join(Separator, values);
    }

    /// <summary>Parses an escaped path. Legacy unescaped separators continue to represent nesting.</summary>
    public static IReadOnlyList<string> Split(string? path) {
        if (string.IsNullOrWhiteSpace(path)) return Array.Empty<string>();
        string value = path!.Trim();
        var parts = new List<string>();
        var current = new StringBuilder();
        for (int index = 0; index < value.Length; index++) {
            char character = value[index];
            if (character == '\\' && index + 1 < value.Length &&
                (value[index + 1] == '\\' || value[index + 1] == '>')) {
                current.Append(value[++index]);
                continue;
            }
            if (character == ' ' &&
                index + Separator.Length <= value.Length &&
                string.CompareOrdinal(value, index, Separator, 0, Separator.Length) == 0) {
                AddPart(parts, current);
                index += Separator.Length - 1;
                continue;
            }
            current.Append(character);
        }
        AddPart(parts, current);
        return parts.Count == 0 ? Array.Empty<string>() : parts.ToArray();
    }

    /// <summary>Returns the unescaped display form of a heading path.</summary>
    public static string? ToDisplayString(string? path) {
        IReadOnlyList<string> parts = Split(path);
        return parts.Count == 0 ? null : string.Join(Separator, parts);
    }

    /// <summary>
    /// Associates an escaped hierarchy-only path with the current public display path.
    /// Reader adapters use this to preserve literal heading delimiters without changing normal output.
    /// </summary>
    public static void SetHierarchyPath(ReaderLocation location, string? hierarchyPath) {
        if (location == null) throw new ArgumentNullException(nameof(location));
        location.HierarchyHeadingPath = hierarchyPath;
        location.HierarchyHeadingDisplayPath = location.HeadingPath;
    }

    internal static string? GetValidatedHierarchyPath(ReaderLocation location) {
        string? hierarchyPath = location.HierarchyHeadingPath;
        if (string.IsNullOrWhiteSpace(hierarchyPath)) return null;
        string? displayPath = location.HierarchyHeadingDisplayPath ?? ToDisplayString(hierarchyPath);
        return string.Equals(displayPath, location.HeadingPath, StringComparison.Ordinal)
            ? hierarchyPath
            : null;
    }

    private static string Escape(string value) => value
        .Replace("\\", "\\\\")
        .Replace(">", "\\>");

    private static void AddPart(ICollection<string> parts, StringBuilder current) {
        string part = current.ToString().Trim();
        current.Clear();
        if (part.Length > 0) parts.Add(part);
    }
}
