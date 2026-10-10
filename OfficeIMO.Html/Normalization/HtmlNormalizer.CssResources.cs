using System.Text.RegularExpressions;

namespace OfficeIMO.Html;

public static partial class HtmlNormalizer {
    private static readonly Regex CssUrlExpression = new Regex("(?<name>(?:[uU]|\\\\0{0,4}(?:75|55)\\s?|\\\\[uU])(?:[rR]|\\\\0{0,4}(?:72|52)\\s?|\\\\[rR])(?:[lL]|\\\\0{0,4}(?:6[cC]|4[cC])\\s?|\\\\[lL]))\\(\\s*(?:\"(?<url>[^\"]*)\"|'(?<url>[^']*)'|(?<url>[^)]+))\\s*\\)", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant | RegexOptions.Compiled);

    private static string NormalizeCssUrls(string css, Uri? baseUri, HtmlUrlPolicy policy, CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        if (string.IsNullOrWhiteSpace(css)) {
            return string.Empty;
        }

        var replacements = new List<CssReplacement>();
        foreach (Match match in CssUrlExpression.Matches(css)) {
            cancellationToken.ThrowIfCancellationRequested();
            if (!IsCssFunctionNameAt(css, match.Index, "url") || IsInsideCssString(css, match.Index)) {
                continue;
            }

            string source = DecodeCssEscapes(match.Groups["url"].Value.Trim().Trim('\'', '"'));
            string resolved = HtmlUrlPolicyEvaluator.ResolveUrl(source, baseUri, policy);
            cancellationToken.ThrowIfCancellationRequested();
            string replacement = string.IsNullOrWhiteSpace(resolved)
                ? "url(\"\")"
                : "url(\"" + EscapeCssString(resolved) + "\")";
            replacements.Add(new CssReplacement(match.Index, match.Index + match.Length, replacement));
        }

        AddCssImportResourceReplacements(css, baseUri, policy, replacements, cancellationToken);
        AddImageSetStringResourceReplacements(css, baseUri, policy, replacements, cancellationToken);
        return ApplyCssReplacements(css, replacements, cancellationToken);
    }

    private static void AddCssImportResourceReplacements(string css, Uri? baseUri, HtmlUrlPolicy policy, ICollection<CssReplacement> replacements, CancellationToken cancellationToken) {
        int index = 0;
        while (index < css.Length) {
            cancellationToken.ThrowIfCancellationRequested();
            if (!TryFindNextCssAtRule(css, index, "import", out int importStart, out int importNameEnd)) {
                return;
            }

            int cursor = SkipCssWhitespaceAndComments(css, importNameEnd);
            if (!TryReadCssImportValue(css, cursor, out int sourceStart, out int sourceEnd)) {
                index = importNameEnd;
                continue;
            }

            string source = DecodeCssEscapes(css.Substring(sourceStart, sourceEnd - sourceStart).Trim());
            string resolved = HtmlUrlPolicyEvaluator.ResolveUrl(source, baseUri, policy);
            cancellationToken.ThrowIfCancellationRequested();
            replacements.Add(new CssReplacement(sourceStart, sourceEnd, EscapeCssString(resolved)));
            index = sourceEnd;
        }
    }

    private static bool TryReadCssImportValue(string css, int cursor, out int sourceStart, out int sourceEnd) {
        sourceStart = 0;
        sourceEnd = 0;
        if (IsCssFunctionNameAt(css, cursor, "url")) {
            int open = css.IndexOf('(', cursor);
            cursor = SkipCssWhitespaceAndComments(css, open + 1);
            if (cursor < css.Length && (css[cursor] == '"' || css[cursor] == '\'')) {
                if (!TryReadCssQuotedValue(css, cursor, out _, out int end)) {
                    return false;
                }

                sourceStart = cursor + 1;
                sourceEnd = end - 1;
                return true;
            }

            sourceStart = cursor;
            while (cursor < css.Length && css[cursor] != ')') {
                cursor++;
            }

            sourceEnd = TrimCssValueEnd(css, sourceStart, cursor);
            return sourceEnd >= sourceStart;
        }

        if (cursor < css.Length && (css[cursor] == '"' || css[cursor] == '\'')) {
            if (!TryReadCssQuotedValue(css, cursor, out _, out int end)) {
                return false;
            }

            sourceStart = cursor + 1;
            sourceEnd = end - 1;
            return true;
        }

        sourceStart = cursor;
        while (cursor < css.Length && !char.IsWhiteSpace(css[cursor]) && css[cursor] != ';') {
            cursor++;
        }

        sourceEnd = cursor;
        return sourceEnd > sourceStart;
    }

    private static int TrimCssValueEnd(string css, int start, int end) {
        while (end > start && char.IsWhiteSpace(css[end - 1])) {
            end--;
        }

        return end;
    }

    private static void AddImageSetStringResourceReplacements(string css, Uri? baseUri, HtmlUrlPolicy policy, ICollection<CssReplacement> replacements, CancellationToken cancellationToken) {
        int index = 0;
        while (index < css.Length) {
            cancellationToken.ThrowIfCancellationRequested();
            if (!TryFindNextCssFunction(css, index, out int functionStart, out int open, "image-set", "-webkit-image-set")) {
                return;
            }

            if (IsInsideCssString(css, functionStart)) {
                index = open + 1;
                continue;
            }

            int close = FindMatchingCssParenthesis(css, open);
            if (close <= open) {
                return;
            }

            int cursor = open + 1;
            while (cursor < close) {
                cancellationToken.ThrowIfCancellationRequested();
                char current = css[cursor];
                if ((current == '"' || current == '\'') && !IsCssTypeFunctionString(css, cursor)) {
                    if (TryReadCssQuotedValue(css, cursor, out string source, out int end)) {
                        string resolved = HtmlUrlPolicyEvaluator.ResolveUrl(DecodeCssEscapes(source), baseUri, policy);
                        cancellationToken.ThrowIfCancellationRequested();
                        replacements.Add(new CssReplacement(cursor + 1, end - 1, EscapeCssString(resolved)));
                        cursor = end;
                        continue;
                    }
                }

                cursor++;
            }

            index = close + 1;
        }
    }

    private static string ApplyCssReplacements(string css, IEnumerable<CssReplacement> replacements, CancellationToken cancellationToken) {
        var ordered = replacements
            .OrderByDescending(range => range.Start)
            .ThenByDescending(range => range.End - range.Start)
            .ToList();
        var builder = new StringBuilder(css);
        var applied = new List<CssReplacement>();
        foreach (CssReplacement replacement in ordered) {
            cancellationToken.ThrowIfCancellationRequested();
            if (applied.Any(appliedReplacement => RangesOverlap(replacement, appliedReplacement))) {
                continue;
            }

            builder.Remove(replacement.Start, replacement.End - replacement.Start);
            builder.Insert(replacement.Start, replacement.Value);
            applied.Add(replacement);
        }

        return builder.ToString();
    }

    private static bool RangesOverlap(CssReplacement first, CssReplacement second) {
        return first.Start < second.End && second.Start < first.End;
    }

    private static string EscapeCssString(string value) {
        return value.Replace("\\", "\\\\").Replace("\"", "\\\"");
    }

    private static bool IsCssFunctionNameAt(string css, int index, string functionName) {
        int open = css.IndexOf('(', index);
        if (open <= index) {
            return false;
        }

        string rawName = css.Substring(index, open - index).Trim();
        if (!string.Equals(DecodeCssEscapes(rawName), functionName, StringComparison.OrdinalIgnoreCase)) {
            return false;
        }

        return index == 0 || !IsCssIdentifierCharacter(css[index - 1]);
    }

    private static bool TryFindNextCssFunction(string css, int startIndex, out int functionStart, out int open, params string[] functionNames) {
        for (open = css.IndexOf('(', Math.Max(0, startIndex)); open >= 0; open = css.IndexOf('(', open + 1)) {
            int nameEnd = open;
            int cursor = nameEnd - 1;
            while (cursor >= 0 && char.IsWhiteSpace(css[cursor])) {
                cursor--;
            }

            int trimmedEnd = cursor + 1;
            while (cursor >= 0 && (IsCssIdentifierCharacter(css[cursor]) || css[cursor] == '\\')) {
                cursor--;
            }

            int nameStart = cursor + 1;
            if (nameStart >= trimmedEnd || (nameStart > 0 && IsCssIdentifierCharacter(css[nameStart - 1]))) {
                continue;
            }

            string decodedName = DecodeCssEscapes(css.Substring(nameStart, trimmedEnd - nameStart));
            foreach (string functionName in functionNames) {
                if (string.Equals(decodedName, functionName, StringComparison.OrdinalIgnoreCase)) {
                    functionStart = nameStart;
                    return true;
                }
            }
        }

        functionStart = -1;
        open = -1;
        return false;
    }

    private static bool IsCssTypeFunctionString(string css, int quoteIndex) {
        int cursor = quoteIndex - 1;
        cursor = SkipCssWhitespaceAndCommentsBackward(css, cursor);

        if (cursor < 0 || css[cursor] != '(') {
            return false;
        }

        cursor--;
        cursor = SkipCssWhitespaceAndCommentsBackward(css, cursor);

        int end = cursor + 1;
        while (cursor >= 0 && (IsCssIdentifierCharacter(css[cursor]) || css[cursor] == '\\')) {
            cursor--;
        }

        string functionName = css.Substring(cursor + 1, end - cursor - 1);
        return string.Equals(DecodeCssEscapes(functionName), "type", StringComparison.OrdinalIgnoreCase);
    }

    private static int SkipCssWhitespaceAndCommentsBackward(string css, int cursor) {
        while (cursor >= 0) {
            if (char.IsWhiteSpace(css[cursor])) {
                cursor--;
                continue;
            }

            if (cursor > 0 && css[cursor - 1] == '*' && css[cursor] == '/') {
                int commentStart = css.LastIndexOf("/*", cursor - 2, StringComparison.Ordinal);
                if (commentStart < 0) {
                    return cursor;
                }

                cursor = commentStart - 1;
                continue;
            }

            break;
        }

        return cursor;
    }

    private static int FindMatchingCssParenthesis(string css, int open) {
        int depth = 0;
        char quote = '\0';
        for (int i = open; i < css.Length; i++) {
            char current = css[i];
            if (quote != '\0') {
                if (current == quote && !IsEscaped(css, i)) {
                    quote = '\0';
                }

                continue;
            }

            if (current == '"' || current == '\'') {
                quote = current;
                continue;
            }

            if (current == '(') {
                depth++;
                continue;
            }

            if (current == ')') {
                depth--;
                if (depth == 0) {
                    return i;
                }
            }
        }

        return -1;
    }

    private static bool TryReadCssQuotedValue(string css, int cursor, out string value, out int end) {
        char quote = css[cursor];
        int start = cursor + 1;
        cursor = start;
        while (cursor < css.Length) {
            if (css[cursor] == quote && !IsEscaped(css, cursor)) {
                value = css.Substring(start, cursor - start);
                end = cursor + 1;
                return true;
            }

            cursor++;
        }

        value = string.Empty;
        end = cursor;
        return false;
    }

    private static bool StartsWith(string text, int index, string value) {
        return index >= 0
            && index + value.Length <= text.Length
            && string.Compare(text, index, value, 0, value.Length, StringComparison.OrdinalIgnoreCase) == 0;
    }

    private static bool TryFindNextCssAtRule(
        string css,
        int startIndex,
        string expectedName,
        out int atRuleStart,
        out int nameEnd) {
        for (int index = css.IndexOf('@', Math.Max(0, startIndex));
             index >= 0;
             index = css.IndexOf('@', index + 1)) {
            if (IsInsideCssString(css, index)) continue;
            if (!TryReadCssAtRuleName(css, index, out string name, out int candidateEnd)) continue;
            if (!string.Equals(name, expectedName, StringComparison.OrdinalIgnoreCase)) continue;
            atRuleStart = index;
            nameEnd = candidateEnd;
            return true;
        }
        atRuleStart = -1;
        nameEnd = -1;
        return false;
    }

    private static bool TryReadCssAtRuleName(string css, int atRuleStart, out string name, out int nameEnd) {
        name = string.Empty;
        nameEnd = atRuleStart;
        if (atRuleStart < 0 || atRuleStart >= css.Length || css[atRuleStart] != '@') return false;
        int cursor = atRuleStart + 1;
        while (cursor < css.Length) {
            if (IsCssIdentifierCharacter(css[cursor])) {
                cursor++;
                continue;
            }
            if (css[cursor] != '\\'
                || !HtmlCssEscapeDecoder.TryDecodeEscape(css, cursor, out _, out int consumed)
                || consumed <= 1) {
                break;
            }
            cursor += consumed;
        }
        if (cursor == atRuleStart + 1) return false;
        name = DecodeCssEscapes(css.Substring(atRuleStart + 1, cursor - atRuleStart - 1));
        nameEnd = cursor;
        return name.Length > 0;
    }

    private static int SkipCssWhitespaceAndComments(string css, int index) {
        while (index < css.Length) {
            if (char.IsWhiteSpace(css[index])) {
                index++;
                continue;
            }

            if (index + 1 < css.Length && css[index] == '/' && css[index + 1] == '*') {
                int commentEnd = css.IndexOf("*/", index + 2, StringComparison.Ordinal);
                if (commentEnd < 0) {
                    return css.Length;
                }

                index = commentEnd + 2;
                continue;
            }

            break;
        }

        return index;
    }

    private static bool IsCssIdentifierCharacter(char value) {
        return char.IsLetterOrDigit(value)
            || value == '_'
            || value == '-'
            || value >= 0x80;
    }

    private static string DecodeCssEscapes(string source) {
        return HtmlCssEscapeDecoder.Decode(source);
    }

    private static bool IsInsideCssString(string css, int index) {
        char quote = '\0';
        for (int i = 0; i < index && i < css.Length; i++) {
            char current = css[i];
            if (quote != '\0') {
                if (current == quote && !IsEscaped(css, i)) {
                    quote = '\0';
                }

                continue;
            }

            if (current == '"' || current == '\'') {
                quote = current;
            }
        }

        return quote != '\0';
    }

    private static bool IsEscaped(string text, int index) {
        int slashCount = 0;
        for (int i = index - 1; i >= 0 && text[i] == '\\'; i--) {
            slashCount++;
        }

        return slashCount % 2 == 1;
    }

    private static string EscapeRawTextElementContent(string value, string elementName) {
        return Regex.Replace(
            value,
            "</\\s*" + Regex.Escape(elementName),
            match => "<\\/" + match.Value.Substring(2),
            RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
    }

    private sealed class CssReplacement {
        internal CssReplacement(int start, int end, string value) {
            Start = start;
            End = end;
            Value = value;
        }

        internal int Start { get; }
        internal int End { get; }
        internal string Value { get; }
    }
}
