namespace OfficeIMO.Markdown;

using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;

/// <summary>
/// Helpers for applying markdown transforms while preserving inline-code spans verbatim.
/// </summary>
public static class MarkdownInlineCode {
    private static readonly Regex InlineCodeSpanRegex = new Regex(
        @"`[^`\r\n]*`",
        RegexOptions.Compiled | RegexOptions.CultureInvariant);

    private static readonly Regex UnmatchedInlineCodeTailRegex = new Regex(
        @"`[^\r\n]*$",
        RegexOptions.Compiled | RegexOptions.CultureInvariant | RegexOptions.Multiline);

    private static readonly Regex InlineCodePlaceholderRegex = new Regex(
        "\u001FOMDCODE_(?<prefix>[0-9a-f]{8})_(?<index>\\d+)\u001E",
        RegexOptions.Compiled | RegexOptions.CultureInvariant);

    /// <summary>
    /// Applies <paramref name="transformer"/> to <paramref name="input"/> while temporarily replacing inline-code
    /// spans and unmatched inline-code tails with unique placeholders.
    /// </summary>
    /// <param name="input">Markdown text to transform.</param>
    /// <param name="transformer">Transformation to apply to the protected text.</param>
    /// <returns>The transformed text with original inline-code spans restored.</returns>
    public static string ApplyTransformPreservingInlineCodeSpans(string? input, Func<string, string> transformer) {
        if (transformer == null) {
            throw new ArgumentNullException(nameof(transformer));
        }

        var protectedInput = ProtectInlineCodeSpans(input, out var codeSpans, out var tokenPrefix);
        var transformed = transformer(protectedInput);
        return RestoreInlineCodeSpans(transformed, codeSpans, tokenPrefix);
    }

    private static string ProtectInlineCodeSpans(string? input, out List<string> codeSpans, out string tokenPrefix) {
        var capturedCodeSpans = new List<string>();
        var value = input ?? string.Empty;
        if (value.Length == 0 || value.IndexOf('`') < 0) {
            codeSpans = capturedCodeSpans;
            tokenPrefix = string.Empty;
            return value;
        }

        var guidText = Guid.NewGuid().ToString("N");
        var prefix = "\u001FOMDCODE_" + guidText.Substring(0, 8) + "_";

        var protectedInput = InlineCodeSpanRegex.Replace(value, match => {
            var index = capturedCodeSpans.Count;
            capturedCodeSpans.Add(match.Value);
            return prefix + index.ToString() + "\u001E";
        });

        protectedInput = UnmatchedInlineCodeTailRegex.Replace(protectedInput, match => {
            var index = capturedCodeSpans.Count;
            capturedCodeSpans.Add(match.Value);
            return prefix + index.ToString() + "\u001E";
        });

        codeSpans = capturedCodeSpans;
        tokenPrefix = prefix;
        return protectedInput;
    }

    private static string RestoreInlineCodeSpans(string input, IReadOnlyList<string> codeSpans, string tokenPrefix) {
        if (string.IsNullOrEmpty(input) || codeSpans.Count == 0 || string.IsNullOrEmpty(tokenPrefix)) {
            return input ?? string.Empty;
        }

        return InlineCodePlaceholderRegex.Replace(input, match => {
            if (!match.Value.StartsWith(tokenPrefix, StringComparison.Ordinal)) {
                return match.Value;
            }

            if (!int.TryParse(match.Groups["index"].Value, out var index)) {
                return match.Value;
            }

            return index >= 0 && index < codeSpans.Count
                ? codeSpans[index]
                : match.Value;
        });
    }
}
