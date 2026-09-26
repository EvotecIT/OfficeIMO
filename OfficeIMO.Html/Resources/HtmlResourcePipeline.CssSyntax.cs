using System.Runtime.CompilerServices;
using System.Text;
using System.Text.RegularExpressions;

namespace OfficeIMO.Html;

public static partial class HtmlResourcePipeline {
    private static readonly ConditionalWeakTable<string, CssStructuralIndex> StructuralIndexes = new();
    private static readonly ConditionalWeakTable<string, CssStringIndex> StringIndexes = new();
    private static readonly ConditionalWeakTable<string, CssAtRuleIndex> AtRuleIndexes = new();

    private static int GetDeclarationStart(string css, int index) {
        FindPreviousCssStructuralTokens(css, index, out int blockStart, out _, out int previousStatementEnd);
        return Math.Max(0, Math.Max(blockStart, previousStatementEnd) + 1);
    }

    private static void FindPreviousCssStructuralTokens(
        string css,
        int beforeIndex,
        out int blockStart,
        out int blockEnd,
        out int statementEnd) {
        CssStructuralIndex index = StructuralIndexes.GetValue(css, BuildStructuralIndex);
        int limit = Math.Min(Math.Max(beforeIndex, 0), css.Length);
        blockStart = index.LastBlockStartBefore(limit);
        blockEnd = index.LastBlockEndBefore(limit);
        statementEnd = index.LastStatementEndBefore(limit);
    }

    private static CssStructuralIndex BuildStructuralIndex(string css) {
        var blockStarts = new List<int>();
        var blockEnds = new List<int>();
        var statementEnds = new List<int>();
        char quote = '\0';
        int parenthesesDepth = 0;
        for (int index = 0; index < css.Length; index++) {
            char current = css[index];
            if (quote != '\0') {
                if (current == quote && !IsEscaped(css, index)) quote = '\0';
                continue;
            }
            if (current is '\'' or '"') { quote = current; continue; }
            if (current == '/' && index + 1 < css.Length && css[index + 1] == '*') {
                index += 2;
                while (index + 1 < css.Length && !(css[index] == '*' && css[index + 1] == '/')) index++;
                if (index + 1 < css.Length) index++;
                continue;
            }
            if (current == '(') { parenthesesDepth++; continue; }
            if (current == ')' && parenthesesDepth > 0) { parenthesesDepth--; continue; }
            if (parenthesesDepth != 0) continue;
            if (current == '{') blockStarts.Add(index);
            else if (current == '}') blockEnds.Add(index);
            else if (current == ';') statementEnds.Add(index);
        }
        return new CssStructuralIndex(blockStarts.ToArray(), blockEnds.ToArray(), statementEnds.ToArray());
    }

    private sealed class CssStructuralIndex {
        private readonly int[] _blockStarts;
        private readonly int[] _blockEnds;
        private readonly int[] _statementEnds;
        private readonly Dictionary<int, string> _selectors = new();

        internal CssStructuralIndex(int[] blockStarts, int[] blockEnds, int[] statementEnds) {
            _blockStarts = blockStarts;
            _blockEnds = blockEnds;
            _statementEnds = statementEnds;
        }

        internal int LastBlockStartBefore(int limit) => LastBefore(_blockStarts, limit);
        internal int LastBlockEndBefore(int limit) => LastBefore(_blockEnds, limit);
        internal int LastStatementEndBefore(int limit) => LastBefore(_statementEnds, limit);
        internal int NextBlockStartAtOrAfter(int index) => NextAtOrAfter(_blockStarts, index);
        internal int NextBlockEndAtOrAfter(int index) => NextAtOrAfter(_blockEnds, index);
        internal int NextStatementEndAtOrAfter(int index) => NextAtOrAfter(_statementEnds, index);

        internal string GetSelector(string css, int blockStart) {
            lock (_selectors) {
                if (_selectors.TryGetValue(blockStart, out string? selector)) return selector;
                int selectorStart = Math.Max(0,
                    Math.Max(LastBlockEndBefore(blockStart), LastStatementEndBefore(blockStart)) + 1);
                selector = css.Substring(selectorStart, blockStart - selectorStart)
                    .Replace(CssCommentMask.ToString(), string.Empty)
                    .Trim();
                int groupingStart = selector.LastIndexOf('{');
                if (groupingStart >= 0) selector = selector.Substring(groupingStart + 1).Trim();
                _selectors.Add(blockStart, selector);
                return selector;
            }
        }

        private static int LastBefore(int[] positions, int limit) {
            int low = 0;
            int high = positions.Length - 1;
            int found = -1;
            while (low <= high) {
                int middle = low + (high - low) / 2;
                if (positions[middle] < limit) {
                    found = positions[middle];
                    low = middle + 1;
                } else high = middle - 1;
            }
            return found;
        }

        private static int NextAtOrAfter(int[] positions, int index) {
            int low = 0;
            int high = positions.Length;
            while (low < high) {
                int middle = low + (high - low) / 2;
                if (positions[middle] < index) low = middle + 1;
                else high = middle;
            }
            return low < positions.Length ? positions[low] : -1;
        }
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

            if (TryParseImageSetOptions(css, functionStart, open + 1, close, out List<CssStringUrlReference> references)) {
                foreach (CssStringUrlReference reference in references) yield return reference;
            }

            index = close + 1;
        }
    }

    private static bool TryParseImageSetOptions(
        string css,
        int functionStart,
        int contentStart,
        int contentEnd,
        out List<CssStringUrlReference> references) {
        references = new List<CssStringUrlReference>();
        int optionStart = contentStart;
        int depth = 0;
        char quote = '\0';
        for (int cursor = contentStart; cursor <= contentEnd; cursor++) {
            char current = cursor < contentEnd ? css[cursor] : ',';
            if (quote != '\0') {
                if (current == quote && !IsEscaped(css, cursor)) quote = '\0';
                continue;
            }
            if (current is '"' or '\'') { quote = current; continue; }
            if (current == '(') { depth++; continue; }
            if (current == ')' && depth > 0) { depth--; continue; }
            if (current != ',' || depth != 0) continue;
            if (!TryParseImageSetOption(css, functionStart, optionStart, cursor, references)) {
                references.Clear();
                return false;
            }
            optionStart = cursor + 1;
        }
        return references.Count > 0 || SkipCssTrivia(css, contentStart, contentEnd) < contentEnd;
    }

    private static bool TryParseImageSetOption(
        string css,
        int functionStart,
        int start,
        int end,
        ICollection<CssStringUrlReference> references) {
        int cursor = SkipCssTrivia(css, start, end);
        if (cursor >= end) return false;
        CssStringUrlReference? stringReference = null;
        if (css[cursor] is '"' or '\'') {
            int quote = cursor;
            if (!TryReadCssQuotedValue(css, cursor, out string source, out int sourceEnd) || sourceEnd > end ||
                string.IsNullOrWhiteSpace(source)) return false;
            stringReference = new CssStringUrlReference(functionStart, sourceEnd, quote + 1, source);
            cursor = sourceEnd;
        } else {
            int nameStart = cursor;
            while (cursor < end && (IsCssIdentifierCharacter(css[cursor]) || css[cursor] == '-')) cursor++;
            if (cursor == nameStart) return false;
            string functionName = css.Substring(nameStart, cursor - nameStart);
            cursor = SkipCssTrivia(css, cursor, end);
            if (cursor >= end || css[cursor] != '(' || !IsSupportedCssImageFunction(functionName)) return false;
            int close = FindMatchingCssParenthesis(css, cursor);
            if (close < cursor || close >= end) return false;
            cursor = close + 1;
        }

        bool sawResolution = false;
        bool sawType = false;
        while ((cursor = SkipCssTrivia(css, cursor, end)) < end) {
            if (IsCssFunctionNameAt(css, cursor, "type")) {
                if (sawType) return false;
                int open = css.IndexOf('(', cursor);
                int close = open >= 0 ? FindMatchingCssParenthesis(css, open) : -1;
                if (close < open || close >= end) return false;
                int valueStart = SkipCssTrivia(css, open + 1, close);
                if (valueStart >= close || css[valueStart] is not ('"' or '\'') ||
                    !TryReadCssQuotedValue(css, valueStart, out string mediaType, out int valueEnd) ||
                    valueEnd > close || SkipCssTrivia(css, valueEnd, close) != close ||
                    string.IsNullOrWhiteSpace(mediaType)) return false;
                sawType = true;
                cursor = close + 1;
                continue;
            }

            int tokenEnd = cursor;
            while (tokenEnd < end && !IsCssWhitespace(css[tokenEnd]) && css[tokenEnd] != CssCommentMask) tokenEnd++;
            if (sawResolution || !IsValidCssResolution(css.Substring(cursor, tokenEnd - cursor))) return false;
            sawResolution = true;
            cursor = tokenEnd;
        }

        if (stringReference != null) references.Add(stringReference);
        return true;
    }

    private static int SkipCssTrivia(string css, int cursor, int end) {
        while (cursor < end && (IsCssWhitespace(css[cursor]) || css[cursor] == CssCommentMask)) cursor++;
        return cursor;
    }

    private static bool IsSupportedCssImageFunction(string name) =>
        name.Equals("url", StringComparison.OrdinalIgnoreCase) ||
        name.Equals("image", StringComparison.OrdinalIgnoreCase) ||
        name.Equals("image-set", StringComparison.OrdinalIgnoreCase) ||
        name.Equals("cross-fade", StringComparison.OrdinalIgnoreCase) ||
        name.Equals("element", StringComparison.OrdinalIgnoreCase) ||
        name.EndsWith("gradient", StringComparison.OrdinalIgnoreCase) ||
        name.Equals("-webkit-gradient", StringComparison.OrdinalIgnoreCase) ||
        name.Equals("paint", StringComparison.OrdinalIgnoreCase);

    private static bool IsValidCssResolution(string token) {
        string unit;
        if (token.EndsWith("dppx", StringComparison.OrdinalIgnoreCase)) unit = "dppx";
        else if (token.EndsWith("dpcm", StringComparison.OrdinalIgnoreCase)) unit = "dpcm";
        else if (token.EndsWith("dpi", StringComparison.OrdinalIgnoreCase)) unit = "dpi";
        else if (token.EndsWith("x", StringComparison.OrdinalIgnoreCase)) unit = "x";
        else return false;
        string number = token.Substring(0, token.Length - unit.Length);
        return double.TryParse(
                number,
                System.Globalization.NumberStyles.Float,
                System.Globalization.CultureInfo.InvariantCulture,
                out double value) &&
            value > 0D && !double.IsNaN(value) && !double.IsInfinity(value);
    }

    private static bool IsCssWhitespace(char value) => value is '\t' or '\n' or '\f' or '\r' or ' ';

    private static bool IsValidCssUrlMatch(string css, Match match) {
        Group source = match.Groups["url"];
        if (!source.Success) return false;
        int open = css.IndexOf('(', match.Index, source.Index - match.Index + 1);
        if (open < 0) return false;
        int tokenStart = open + 1;
        while (tokenStart < source.Index && IsCssWhitespace(css[tokenStart])) tokenStart++;
        bool quoted = tokenStart < css.Length && (css[tokenStart] == '\'' || css[tokenStart] == '"');
        int start = source.Index;
        int end = source.Index + source.Length;
        if (!quoted) {
            while (start < end && IsCssWhitespace(css[start])) start++;
            while (end > start && IsCssWhitespace(css[end - 1])) end--;
        }
        if (start == end) return false;
        for (int index = start; index < end; index++) {
            char value = css[index];
            if (value == '\\') {
                if (++index >= end) return false;
                if (css[index] is '\r' or '\n' or '\f') {
                    if (!quoted) return false;
                    if (css[index] == '\r' && index + 1 < end && css[index + 1] == '\n') index++;
                    continue;
                }
                if (IsCssHexDigit(css[index])) {
                    int digits = 1;
                    while (digits < 6 && index + 1 < end && IsCssHexDigit(css[index + 1])) {
                        index++;
                        digits++;
                    }
                    if (index + 1 < end && IsCssWhitespace(css[index + 1])) {
                        index++;
                        if (css[index] == '\r' && index + 1 < end && css[index + 1] == '\n') index++;
                    }
                }
                continue;
            }
            if (quoted) {
                if (value is '\r' or '\n' or '\f') return false;
                continue;
            }
            if (IsCssWhitespace(value) || value is '\'' or '"' or '(' ||
                value <= '\u0008' || value == '\u000B' || value is >= '\u000E' and <= '\u001F' || value == '\u007F') {
                return false;
            }
        }
        return true;
    }

    private static bool IsCssHexDigit(char value) =>
        value is >= '0' and <= '9' or >= 'a' and <= 'f' or >= 'A' and <= 'F';

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
            if (!TryFindNextAtRule(css, index, "import", out int importStart, out int importNameEnd)) {
                yield break;
            }

            if (HasDisallowedRuleBeforeImport(css, importStart)) {
                yield break;
            }

            int cursor = SkipWhitespace(css, importNameEnd);
            string source;
            int end;
            if (IsCssFunctionNameAt(css, cursor, "url")) {
                int open = css.IndexOf('(', cursor);
                cursor = SkipWhitespace(css, open + 1);
                if (!TryReadCssUrlFunctionSource(css, cursor, out source, out end)) {
                    index = importNameEnd;
                    continue;
                }
            } else if (cursor < css.Length && (css[cursor] == '"' || css[cursor] == '\'')) {
                if (!TryReadCssQuotedValue(css, cursor, out source, out end)) {
                    index = importNameEnd;
                    continue;
                }
            } else {
                int sourceStart = cursor;
                while (cursor < css.Length && !IsCssWhitespace(css[cursor]) && css[cursor] != ';') {
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

    private static bool IsApplicableCssImport(
        string conditionText,
        HtmlResourcePipelineOptions options,
        out bool hasUnknownSupportsCondition,
        out bool hasUnknownMediaCondition) {
        hasUnknownSupportsCondition = false;
        hasUnknownMediaCondition = false;
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
                bool isKnown = HtmlComputedStyleEngine.TryEvaluateSupports(supportsCondition, out bool isApplicable);
                hasUnknownSupportsCondition |= !isKnown;
                if (!isKnown) isApplicable = HtmlComputedStyleEngine.IsApplicableSupports(supportsCondition);
                if (!isApplicable) {
                    return false;
                }

                remaining = afterSupports.TrimStart();
                continue;
            }

            break;
        }

        if (remaining.Length == 0) return true;
        hasUnknownMediaCondition = HtmlComputedStyleEngine.HasUnknownMediaFeature(remaining);
        return IsApplicableMedia(remaining, options);
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
        while (index < text.Length && IsCssWhitespace(text[index])) {
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
            while (cursor >= 0 && IsCssWhitespace(css[cursor])) {
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
        return StringIndexes.GetValue(css, BuildStringIndex).Contains(index);
    }

    private static CssStringIndex BuildStringIndex(string css) {
        var opens = new List<int>();
        var closes = new List<int>();
        char quote = '\0';
        for (int i = 0; i < css.Length; i++) {
            char current = css[i];
            if (quote != '\0') {
                if (current == quote && !IsEscaped(css, i)) {
                    closes.Add(i);
                    quote = '\0';
                }
                continue;
            }
            if (current == '"' || current == '\'') {
                quote = current;
                opens.Add(i);
            }
        }
        if (quote != '\0') closes.Add(css.Length);
        return new CssStringIndex(opens.ToArray(), closes.ToArray());
    }

    private sealed class CssStringIndex {
        private readonly int[] _opens;
        private readonly int[] _closes;

        internal CssStringIndex(int[] opens, int[] closes) {
            _opens = opens;
            _closes = closes;
        }

        internal bool Contains(int index) {
            int low = 0;
            int high = _opens.Length - 1;
            int previous = -1;
            while (low <= high) {
                int middle = low + (high - low) / 2;
                if (_opens[middle] < index) {
                    previous = middle;
                    low = middle + 1;
                } else high = middle - 1;
            }
            return previous >= 0 && index <= _closes[previous];
        }
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
        CssStructuralIndex structure = StructuralIndexes.GetValue(css, BuildStructuralIndex);
        int previousBoundary = Math.Max(structure.LastStatementEndBefore(index), structure.LastBlockEndBefore(index));
        return AtRuleIndexes.GetValue(css, BuildAtRuleIndex).HasImportBetween(previousBoundary, index);
    }

    private static bool IsAtRulePreludeUrl(string css, int index) {
        CssStructuralIndex structure = StructuralIndexes.GetValue(css, BuildStructuralIndex);
        int previousBoundary = Math.Max(structure.LastBlockStartBefore(index),
            Math.Max(structure.LastBlockEndBefore(index), structure.LastStatementEndBefore(index)));
        if (!AtRuleIndexes.GetValue(css, BuildAtRuleIndex).HasAtRuleBetween(previousBoundary, index)) return false;
        int nextOpen = structure.NextBlockStartAtOrAfter(index);
        if (nextOpen < 0) return false;
        int nextSemicolon = structure.NextStatementEndAtOrAfter(index);
        int nextClose = structure.NextBlockEndAtOrAfter(index);
        return (nextSemicolon < 0 || nextOpen < nextSemicolon)
            && (nextClose < 0 || nextOpen < nextClose);
    }

    private static CssAtRuleIndex BuildAtRuleIndex(string css) {
        var atRules = new List<int>();
        var imports = new List<int>();
        var fontFaces = new List<int>();
        for (int index = css.IndexOf('@'); index >= 0; index = css.IndexOf('@', index + 1)) {
            if (IsInsideCssString(css, index)) continue;
            atRules.Add(index);
            if (!TryReadAtRuleName(css, index, out string name, out _)) continue;
            if (name.Equals("import", StringComparison.OrdinalIgnoreCase)) imports.Add(index);
            else if (name.Equals("font-face", StringComparison.OrdinalIgnoreCase)) fontFaces.Add(index);
        }
        return new CssAtRuleIndex(atRules.ToArray(), imports.ToArray(), fontFaces.ToArray());
    }

    private sealed class CssAtRuleIndex {
        private readonly int[] _atRules;
        private readonly int[] _imports;
        private readonly int[] _fontFaces;

        internal CssAtRuleIndex(int[] atRules, int[] imports, int[] fontFaces) {
            _atRules = atRules;
            _imports = imports;
            _fontFaces = fontFaces;
        }

        internal bool HasAtRuleBetween(int previousBoundary, int index) => HasBetween(_atRules, previousBoundary, index);
        internal bool HasImportBetween(int previousBoundary, int index) => HasBetween(_imports, previousBoundary, index);
        internal int LastFontFaceBefore(int index) => LastBefore(_fontFaces, index);

        private static bool HasBetween(int[] positions, int previousBoundary, int index) =>
            LastBefore(positions, index) > previousBoundary;

        private static int LastBefore(int[] positions, int index) {
            int low = 0;
            int high = positions.Length - 1;
            int found = -1;
            while (low <= high) {
                int middle = low + (high - low) / 2;
                if (positions[middle] < index) {
                    found = positions[middle];
                    low = middle + 1;
                } else high = middle - 1;
            }
            return found;
        }
    }

    private static bool HasDisallowedRuleBeforeImport(string css, int index) {
        int cursor = 0;
        while (cursor < index) {
            cursor = SkipWhitespace(css, cursor);
            if (cursor >= index) return false;
            if (StartsWith(css, cursor, "<!--")) {
                cursor += 4;
                continue;
            }
            if (StartsWith(css, cursor, "-->")) {
                cursor += 3;
                continue;
            }
            bool allowed = IsAtRuleAt(css, cursor, "@charset")
                || IsAtRuleAt(css, cursor, "@import")
                || IsAtRuleAt(css, cursor, "@layer");
            if (!allowed) return true;

            char quote = '\0';
            int parentheses = 0;
            bool terminated = false;
            for (; cursor < index; cursor++) {
                char current = css[cursor];
                if (quote != '\0') {
                    if (current == quote && !IsEscaped(css, cursor)) quote = '\0';
                    continue;
                }
                if (current == '"' || current == '\'') {
                    quote = current;
                    continue;
                }
                if (current == '(') {
                    parentheses++;
                    continue;
                }
                if (current == ')' && parentheses > 0) {
                    parentheses--;
                    continue;
                }
                if (parentheses == 0 && current == '{') return true;
                if (parentheses == 0 && current == ';') {
                    cursor++;
                    terminated = true;
                    break;
                }
            }
            if (!terminated) return true;
        }
        return false;
    }

    private static bool IsAtRuleAt(string css, int index, string rule) {
        if (string.IsNullOrEmpty(rule) || rule[0] != '@') return false;
        return TryReadAtRuleName(css, index, out string name, out _)
            && string.Equals(name, rule.Substring(1), StringComparison.OrdinalIgnoreCase);
    }

    private static bool TryFindNextAtRule(
        string css,
        int startIndex,
        string expectedName,
        out int atRuleStart,
        out int nameEnd) {
        for (int index = css.IndexOf('@', Math.Max(0, startIndex));
             index >= 0;
             index = css.IndexOf('@', index + 1)) {
            if (IsInsideCssString(css, index)) continue;
            if (!TryReadAtRuleName(css, index, out string name, out int candidateEnd)) continue;
            if (!string.Equals(name, expectedName, StringComparison.OrdinalIgnoreCase)) continue;
            atRuleStart = index;
            nameEnd = candidateEnd;
            return true;
        }

        atRuleStart = -1;
        nameEnd = -1;
        return false;
    }

    private static bool TryReadAtRuleName(string css, int atRuleStart, out string name, out int nameEnd) {
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

    private static HtmlResourceKind ClassifyCssUrl(string css, int index) {
        string propertyName = GetCssDeclarationPropertyName(css, index);
        CssStructuralIndex structure = StructuralIndexes.GetValue(css, BuildStructuralIndex);
        int blockStart = structure.LastBlockStartBefore(index);
        int fontFaceStart = AtRuleIndexes.GetValue(css, BuildAtRuleIndex).LastFontFaceBefore(blockStart);
        int previousBlockEnd = structure.LastBlockEndBefore(blockStart);
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

    internal static bool IsSupportedCssImageUrlProperty(string propertyName) {
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
            case "shape-outside":
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
        return HtmlCssEscapeDecoder.Decode(source.Replace(CssCommentMask.ToString(), string.Empty));
    }

}
