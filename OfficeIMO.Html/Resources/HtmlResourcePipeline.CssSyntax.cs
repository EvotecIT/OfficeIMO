using System.Text;
using System.Text.RegularExpressions;

namespace OfficeIMO.Html;

public static partial class HtmlResourcePipeline {
    private static int GetDeclarationStart(string css, int index) {
        int blockStart = css.LastIndexOf('{', Math.Max(0, index - 1));
        int previousStatementEnd = css.LastIndexOf(';', Math.Max(0, index - 1));
        return Math.Max(0, Math.Max(blockStart, previousStatementEnd) + 1);
    }

    private static IEnumerable<CssStringUrlReference> ExtractImageSetStringUrls(string css) {
        int index = 0;
        while (index < css.Length) {
            if (!TryFindNextCssFunction(css, index, out int functionStart, out int open, "image-set", "-webkit-image-set")) {
                yield break;
            }

            if (IsInsideCssString(css, functionStart)) {
                index = open + 1;
                continue;
            }

            int close = FindMatchingCssParenthesis(css, open);
            if (close <= open) {
                yield break;
            }

            int valueCursor = open + 1;
            while (valueCursor < close) {
                char current = css[valueCursor];
                if ((current == '"' || current == '\'') && !IsCssTypeFunctionString(css, valueCursor)) {
                    if (TryReadCssQuotedValue(css, valueCursor, out string source, out int end)) {
                        if (!string.IsNullOrWhiteSpace(source)) {
                            yield return new CssStringUrlReference(functionStart, end, source);
                        }

                        valueCursor = end;
                        continue;
                    }
                }

                valueCursor++;
            }

            index = close + 1;
        }
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
        return CssFunctionNameEquals(functionName, "type");
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

    private static IEnumerable<CssImportReference> ExtractCssImports(string css) {
        int index = 0;
        while (index < css.Length) {
            int importStart = css.IndexOf("@import", index, StringComparison.OrdinalIgnoreCase);
            if (importStart < 0) {
                yield break;
            }

            if (IsInsideCssString(css, importStart)) {
                index = importStart + 7;
                continue;
            }

            if (!HasImportTokenBoundary(css, importStart)) {
                index = importStart + 7;
                continue;
            }

            if (HasStyleRuleBefore(css, importStart)) {
                yield break;
            }

            int cursor = SkipWhitespace(css, importStart + 7);
            string source;
            int end;
            if (IsCssFunctionNameAt(css, cursor, "url")) {
                int open = css.IndexOf('(', cursor);
                cursor = SkipWhitespace(css, open + 1);
                if (!TryReadCssUrlFunctionSource(css, cursor, out source, out end)) {
                    index = importStart + 7;
                    continue;
                }
            } else if (cursor < css.Length && (css[cursor] == '"' || css[cursor] == '\'')) {
                if (!TryReadCssQuotedValue(css, cursor, out source, out end)) {
                    index = importStart + 7;
                    continue;
                }
            } else {
                int sourceStart = cursor;
                while (cursor < css.Length && !char.IsWhiteSpace(css[cursor]) && css[cursor] != ';') {
                    cursor++;
                }

                source = css.Substring(sourceStart, cursor - sourceStart);
                end = cursor;
            }

            int importEnd = end;
            while (importEnd < css.Length && css[importEnd] != ';') {
                importEnd++;
            }

            if (importEnd < css.Length) {
                importEnd++;
            }

            string conditionText = css.Substring(end, Math.Max(0, importEnd - end)).Trim().TrimEnd(';').Trim();
            yield return new CssImportReference(importStart, importEnd, source, conditionText);
            index = importEnd;
        }
    }

    private static bool IsApplicableCssImport(string conditionText, HtmlCssMediaContext mediaContext) {
        string remaining = conditionText.Trim();
        if (remaining.Length == 0) {
            return true;
        }

        while (remaining.Length > 0) {
            if (TryConsumeCssImportFunctionCondition(remaining, "layer", out _, out string afterLayer)) {
                remaining = afterLayer.TrimStart();
                continue;
            }

            if (StartsWithCssIdentifier(remaining, "layer")) {
                remaining = remaining.Substring("layer".Length).TrimStart();
                continue;
            }

            if (TryConsumeCssImportFunctionCondition(remaining, "supports", out string supportsCondition, out string afterSupports)) {
                if (!HtmlComputedStyleEngine.IsApplicableSupports(supportsCondition)) {
                    return false;
                }

                remaining = afterSupports.TrimStart();
                continue;
            }

            break;
        }

        return remaining.Length == 0 || IsApplicableMedia(remaining, mediaContext);
    }

    private static bool TryConsumeCssImportFunctionCondition(string text, string functionName, out string argument, out string remaining) {
        argument = string.Empty;
        remaining = text;
        if (!IsCssFunctionNameAt(text, 0, functionName)) {
            return false;
        }

        int open = text.IndexOf('(');
        if (open < 0) {
            return false;
        }

        int close = FindMatchingCssParenthesis(text, open);
        if (close <= open) {
            return false;
        }

        argument = text.Substring(open + 1, close - open - 1).Trim();
        remaining = text.Substring(close + 1);
        return true;
    }

    private static bool StartsWithCssIdentifier(string text, string identifier) {
        if (!text.StartsWith(identifier, StringComparison.OrdinalIgnoreCase)) {
            return false;
        }

        return text.Length == identifier.Length || !IsCssIdentifierCharacter(text[identifier.Length]);
    }

    private static bool TryReadCssUrlFunctionSource(string css, int cursor, out string source, out int end) {
        if (cursor < css.Length && (css[cursor] == '"' || css[cursor] == '\'')) {
            if (!TryReadCssQuotedValue(css, cursor, out source, out cursor)) {
                end = cursor;
                return false;
            }
        } else {
            int sourceStart = cursor;
            while (cursor < css.Length && css[cursor] != ')') {
                cursor++;
            }

            source = css.Substring(sourceStart, cursor - sourceStart).Trim();
        }

        cursor = SkipWhitespace(css, cursor);
        if (cursor < css.Length && css[cursor] == ')') {
            cursor++;
        }

        end = cursor;
        return true;
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

    private static int SkipWhitespace(string text, int index) {
        while (index < text.Length && char.IsWhiteSpace(text[index])) {
            index++;
        }

        return index;
    }

    private static bool StartsWith(string text, int index, string value) {
        return index >= 0
            && index + value.Length <= text.Length
            && string.Compare(text, index, value, 0, value.Length, StringComparison.OrdinalIgnoreCase) == 0;
    }

    private static bool IsCssFunctionNameAt(string css, int index, string functionName) {
        int open = css.IndexOf('(', index);
        if (open <= index) {
            return false;
        }

        string rawName = css.Substring(index, open - index).Trim();
        if (!CssFunctionNameEquals(rawName, functionName)) {
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

            string rawName = css.Substring(nameStart, trimmedEnd - nameStart);
            foreach (string functionName in functionNames) {
                if (CssFunctionNameEquals(rawName, functionName)) {
                    functionStart = nameStart;
                    return true;
                }
            }
        }

        functionStart = -1;
        open = -1;
        return false;
    }

    private static bool IsCssIdentifierCharacter(char value) {
        return char.IsLetterOrDigit(value)
            || value == '_'
            || value == '-'
            || value >= 0x80;
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

    private static string StripCssCommentsOutsideStrings(string css) {
        var result = new System.Text.StringBuilder(css.Length);
        char quote = '\0';
        for (int i = 0; i < css.Length; i++) {
            char current = css[i];
            if (quote != '\0') {
                result.Append(current);
                if (current == quote && !IsEscaped(css, i)) {
                    quote = '\0';
                }

                continue;
            }

            if (current == '"' || current == '\'') {
                quote = current;
                result.Append(current);
                continue;
            }

            if (current == '/' && i + 1 < css.Length && css[i + 1] == '*') {
                i += 2;
                while (i + 1 < css.Length && !(css[i] == '*' && css[i + 1] == '/')) {
                    i++;
                }

                if (i + 1 < css.Length) {
                    i++;
                }

                result.Append(' ');
                continue;
            }

            result.Append(current);
        }

        return result.ToString();
    }

    private static bool IsCustomPropertyUrl(string css, int index) {
        return TryGetCustomPropertyName(css, index, out _);
    }

    private static bool TryGetCustomPropertyName(string css, int index, out string propertyName) {
        int blockStart = css.LastIndexOf('{', Math.Max(0, index - 1));
        int previousBoundary = Math.Max(css.LastIndexOf(';', Math.Max(0, index - 1)), blockStart);
        string declaration = css.Substring(Math.Max(0, previousBoundary + 1), index - Math.Max(0, previousBoundary + 1)).TrimStart();
        if (!declaration.StartsWith("--", StringComparison.Ordinal)) {
            propertyName = string.Empty;
            return false;
        }

        int separator = declaration.IndexOf(':');
        if (separator <= 0) {
            propertyName = string.Empty;
            return false;
        }

        propertyName = declaration.Substring(0, separator).Trim();
        return propertyName.Length > 2;
    }

    private static bool IsImportAtRuleUrl(string css, int index) {
        int previousSemicolon = css.LastIndexOf(';', Math.Max(0, index - 1));
        int previousBlockEnd = css.LastIndexOf('}', Math.Max(0, index - 1));
        int previousBoundary = Math.Max(previousSemicolon, previousBlockEnd);
        string statement = css.Substring(Math.Max(0, previousBoundary + 1), index - Math.Max(0, previousBoundary + 1));
        int importStart = statement.IndexOf("@import", StringComparison.OrdinalIgnoreCase);
        return importStart >= 0 && HasImportTokenBoundary(statement, importStart);
    }

    private static bool IsAtRulePreludeUrl(string css, int index) {
        int previousOpen = css.LastIndexOf('{', Math.Max(0, index - 1));
        int previousClose = css.LastIndexOf('}', Math.Max(0, index - 1));
        int previousSemicolon = css.LastIndexOf(';', Math.Max(0, index - 1));
        int previousBoundary = Math.Max(previousOpen, Math.Max(previousClose, previousSemicolon));
        int segmentStart = Math.Max(0, previousBoundary + 1);
        string prefix = css.Substring(segmentStart, index - segmentStart);
        if (prefix.LastIndexOf('@') < 0) {
            return false;
        }

        int nextOpen = css.IndexOf('{', index);
        if (nextOpen < 0) {
            return false;
        }

        int nextSemicolon = css.IndexOf(';', index);
        int nextClose = css.IndexOf('}', index);
        return (nextSemicolon < 0 || nextOpen < nextSemicolon)
            && (nextClose < 0 || nextOpen < nextClose);
    }

    private static bool HasImportTokenBoundary(string css, int importStart) {
        return HasAtRuleTokenBoundary(css, importStart, "@import");
    }

    private static bool HasAtRuleTokenBoundary(string css, int atRuleStart, string atRuleName) {
        int afterImport = atRuleStart + atRuleName.Length;
        return afterImport >= css.Length || !IsCssIdentifierCharacter(css[afterImport]);
    }

    private static bool HasStyleRuleBefore(string css, int index) {
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
                continue;
            }

            if (current == '{' || current == '}') {
                return true;
            }
        }

        return false;
    }

    private static HtmlResourceKind ClassifyCssUrl(string css, int index) {
        string propertyName = GetCssDeclarationPropertyName(css, index);
        int blockStart = css.LastIndexOf('{', Math.Max(0, index - 1));
        string blockPrefix = blockStart >= 0 ? css.Substring(0, blockStart).ToLowerInvariant() : string.Empty;
        int fontFaceStart = blockPrefix.LastIndexOf("@font-face", StringComparison.Ordinal);
        int previousBlockEnd = blockPrefix.LastIndexOf('}');
        if (fontFaceStart >= 0 && fontFaceStart > previousBlockEnd) {
            return HtmlResourceKind.Font;
        }

        if (IsSupportedCssImageUrlProperty(propertyName)) {
            return HtmlResourceKind.Image;
        }

        return HtmlResourceKind.Other;
    }

    private static bool IsSupportedCssUrlDeclaration(string css, int index) {
        return ClassifyCssUrl(css, index) != HtmlResourceKind.Other;
    }

    private static string GetCssDeclarationPropertyName(string css, int index) {
        int declarationStart = GetDeclarationStart(css, index);
        int separator = css.IndexOf(':', declarationStart, Math.Max(0, index - declarationStart));
        if (separator <= declarationStart) {
            return string.Empty;
        }

        string propertyName = DecodeCssEscapes(css.Substring(declarationStart, separator - declarationStart).Trim());
        return propertyName.StartsWith("--", StringComparison.Ordinal)
            ? propertyName
            : propertyName.ToLowerInvariant();
    }

    private static bool IsSupportedCssImageUrlProperty(string propertyName) {
        switch (propertyName) {
            case "background":
            case "background-image":
            case "border-image":
            case "border-image-source":
            case "content":
            case "cursor":
            case "list-style":
            case "list-style-image":
            case "mask":
            case "mask-image":
            case "-webkit-mask":
            case "-webkit-mask-image":
            case "filter":
            case "clip-path":
                return true;
            default:
                return false;
        }
    }

    private static bool IsImportUrl(int index, IEnumerable<SourceRange> ranges) {
        return IsInRanges(index, ranges);
    }

    private static bool IsInRanges(int index, IEnumerable<SourceRange> ranges) {
        foreach (SourceRange range in ranges) {
            if (index >= range.Start && index < range.End) {
                return true;
            }
        }

        return false;
    }

    private static string NormalizeSource(string source) {
        return source.Trim().Trim('\'', '"');
    }

    private static string DecodeCssEscapes(string source) {
        return HtmlCssEscapeDecoder.Decode(source);
    }

}
