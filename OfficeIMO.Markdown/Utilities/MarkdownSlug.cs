using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;

namespace OfficeIMO.Markdown;

/// <summary>
/// Slug utilities for generating anchor ids compatible with GitHub-like platforms.
/// </summary>
internal static class MarkdownSlug {
    /// <summary>
    /// Creates a registry that tracks generated slugs to ensure uniqueness within a document render.
    /// </summary>
    public static Dictionary<string, int> CreateRegistry() => new(StringComparer.Ordinal);

    public static string GitHub(string text) => GitHub(text, registry: null);

    /// <summary>
    /// Generates a GitHub-style slug and ensures uniqueness using the provided registry when supplied.
    /// </summary>
    public static string GitHub(string text, IDictionary<string, int>? registry) {
        return Generate(text, MarkdownHeadingIdentifierStyle.OfficeIMO, registry);
    }

    internal static string Generate(string text, MarkdownHeadingIdentifierStyle style, IDictionary<string, int>? registry = null) {
        if (string.IsNullOrEmpty(text)) {
            return EnsureUnique(GetEmptySlugFallback(style), registry);
        }

        var sb = new StringBuilder(text.Length);
        bool prevHyphen = false;
        foreach (char ch in text.ToLowerInvariant()) {
            if (IsAllowedIdentifierCharacter(ch, style)) {
                sb.Append(ch);
                prevHyphen = false;
            } else if (IsWordSeparator(ch, style)) {
                if (!prevHyphen) {
                    sb.Append('-');
                    prevHyphen = true;
                }
            } else {
                // skip punctuation
            }
        }

        var result = sb.ToString().Trim('-');
        if (result.Length == 0) {
            result = GetEmptySlugFallback(style);
        }

        return EnsureUnique(result, registry);
    }

    private static bool IsAllowedIdentifierCharacter(char ch, MarkdownHeadingIdentifierStyle style) {
        if ((ch >= 'a' && ch <= 'z') || (ch >= '0' && ch <= '9')) {
            return true;
        }

        if (style == MarkdownHeadingIdentifierStyle.MarkdigDefault && (ch == '_' || ch == '.')) {
            return true;
        }

        if (style == MarkdownHeadingIdentifierStyle.GitHub && ch == '_') {
            return true;
        }

        return style == MarkdownHeadingIdentifierStyle.GitHub && char.IsLetterOrDigit(ch);
    }

    private static bool IsWordSeparator(char ch, MarkdownHeadingIdentifierStyle style) {
        if (ch == ' ' || ch == '-') {
            return true;
        }

        if (style == MarkdownHeadingIdentifierStyle.OfficeIMO && ch == '_') {
            return true;
        }

        return style == MarkdownHeadingIdentifierStyle.GitHub && char.GetUnicodeCategory(ch) == UnicodeCategory.OtherSymbol;
    }

    private static string GetEmptySlugFallback(MarkdownHeadingIdentifierStyle style) =>
        style == MarkdownHeadingIdentifierStyle.OfficeIMO ? string.Empty : "section";

    private static string EnsureUnique(string slug, IDictionary<string, int>? registry) {
        if (registry is null) return slug;

        if (!registry.TryGetValue(slug, out var count)) {
            registry[slug] = 0;
            return slug;
        }

        int next = count + 1;
        string candidate;
        do {
            candidate = string.IsNullOrEmpty(slug)
                ? $"-{next.ToString(CultureInfo.InvariantCulture)}"
                : $"{slug}-{next.ToString(CultureInfo.InvariantCulture)}";
            if (!registry.ContainsKey(candidate)) {
                registry[slug] = next;
                registry[candidate] = 0;
                return candidate;
            }
            next++;
        } while (true);
    }
}
