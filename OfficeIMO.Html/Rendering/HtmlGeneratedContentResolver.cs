using System.Globalization;
using System.Text;
using AngleSharp.Dom;
using AngleSharp.Html.Dom;

namespace OfficeIMO.Html;

internal static class HtmlGeneratedContentResolver {
    private const string ComponentName = "OfficeIMO.Html.Renderer";

    internal static HtmlGeneratedContentSet Resolve(
        IHtmlDocument document,
        HtmlComputedStyleSet styles,
        HtmlDiagnosticReport diagnostics,
        int maximumDepth) {
        if (maximumDepth <= 0) throw new ArgumentOutOfRangeException(nameof(maximumDepth));
        var content = new Dictionary<IElement, HtmlGeneratedPseudoContentPair>();
        if (!styles.HasPseudoElements) return new HtmlGeneratedContentSet(content);
        var counters = new CounterState();
        IElement? root = document.DocumentElement ?? document.Body;
        if (root != null) {
            int level = counters.EnterLevel();
            TraverseElement(root, level, 0, maximumDepth, styles, diagnostics, counters, content);
            counters.ExitLevel(level);
        }

        return new HtmlGeneratedContentSet(content);
    }

    private static void TraverseElement(
        IElement element,
        int level,
        int depth,
        int maximumDepth,
        HtmlComputedStyleSet styles,
        HtmlDiagnosticReport diagnostics,
        CounterState counters,
        IDictionary<IElement, HtmlGeneratedPseudoContentPair> content) {
        if (depth > maximumDepth) {
            throw new HtmlDomLimitException(
                HtmlRenderDiagnosticCodes.DepthLimitExceeded,
                "HTML generated-content traversal exceeded the configured maximum depth at " + HtmlRenderStyleResolver.DescribeSource(element) + ".",
                nameof(HtmlRenderOptions.MaxLayoutDepth),
                depth,
                maximumDepth);
        }

        if (!styles.Elements.TryGetValue(element, out HtmlComputedStyle? elementStyle)
            || string.Equals(elementStyle.GetValue("display"), "none", StringComparison.OrdinalIgnoreCase)) {
            return;
        }

        ApplyCounterProperties(elementStyle, level, counters, diagnostics, HtmlRenderStyleResolver.DescribeSource(element));
        ResolvePseudo(element, HtmlPseudoElementKind.Before, level, styles, diagnostics, counters, content);

        int childLevel = counters.EnterLevel();
        foreach (IElement child in element.Children) {
            if (!ShouldSkipSubtree(child)) TraverseElement(child, childLevel, depth + 1, maximumDepth, styles, diagnostics, counters, content);
        }

        counters.ExitLevel(childLevel);
        ResolvePseudo(element, HtmlPseudoElementKind.After, level, styles, diagnostics, counters, content);
    }

    private static void ResolvePseudo(
        IElement element,
        HtmlPseudoElementKind kind,
        int level,
        HtmlComputedStyleSet styles,
        HtmlDiagnosticReport diagnostics,
        CounterState counters,
        IDictionary<IElement, HtmlGeneratedPseudoContentPair> content) {
        if (!styles.TryGetPseudoStyle(element, kind, out HtmlComputedStyle pseudoStyle)
            || string.Equals(pseudoStyle.GetValue("display"), "none", StringComparison.OrdinalIgnoreCase)) {
            return;
        }

        string pseudoName = kind == HtmlPseudoElementKind.Before ? "::before" : "::after";
        string pseudoSource = HtmlRenderStyleResolver.DescribeSource(element) + pseudoName;
        ApplyCounterProperties(pseudoStyle, level, counters, diagnostics, pseudoSource);
        string expression = pseudoStyle.GetValue("content");
        if (string.IsNullOrWhiteSpace(expression)
            || string.Equals(expression.Trim(), "none", StringComparison.OrdinalIgnoreCase)
            || string.Equals(expression.Trim(), "normal", StringComparison.OrdinalIgnoreCase)) {
            return;
        }

        if (!TryEvaluate(expression, element, counters, out string generated, out string detail)) {
            diagnostics.Add(
                ComponentName,
                HtmlRenderDiagnosticCodes.GeneratedContentUnsupported,
                "A generated-content expression could not be represented and was omitted.",
                HtmlDiagnosticSeverity.Warning,
                pseudoSource,
                detail.Length > 0 ? detail : expression,
                HtmlConversionLossKind.Omission);
            return;
        }

        if (generated.Length == 0) return;
        if (!content.TryGetValue(element, out HtmlGeneratedPseudoContentPair? pair)) {
            pair = new HtmlGeneratedPseudoContentPair();
            content[element] = pair;
        }

        if (kind == HtmlPseudoElementKind.Before) pair.Before = generated;
        else pair.After = generated;
    }

    private static void ApplyCounterProperties(
        HtmlComputedStyle style,
        int level,
        CounterState counters,
        HtmlDiagnosticReport diagnostics,
        string source) {
        ApplyCounterProperty(style.GetValue("counter-reset"), "counter-reset", 0, level, counters.Reset, diagnostics, source);
        ApplyCounterProperty(style.GetValue("counter-set"), "counter-set", 0, level, counters.Set, diagnostics, source);
        ApplyCounterProperty(style.GetValue("counter-increment"), "counter-increment", 1, level, counters.Increment, diagnostics, source);
    }

    private static void ApplyCounterProperty(
        string value,
        string property,
        int defaultValue,
        int level,
        Action<string, int, int> apply,
        HtmlDiagnosticReport diagnostics,
        string source) {
        if (!TryParseCounterOperations(value, defaultValue, out IReadOnlyList<CounterOperation> operations)) {
            diagnostics.Add(
                ComponentName,
                HtmlRenderDiagnosticCodes.GeneratedCounterUnsupported,
                "A CSS counter declaration could not be represented and was ignored.",
                HtmlDiagnosticSeverity.Warning,
                source,
                property + "=" + value,
                HtmlConversionLossKind.Omission);
            return;
        }

        foreach (CounterOperation operation in operations) apply(operation.Name, operation.Value, level);
    }

    private static bool TryParseCounterOperations(string value, int defaultValue, out IReadOnlyList<CounterOperation> operations) {
        if (string.IsNullOrWhiteSpace(value) || string.Equals(value.Trim(), "none", StringComparison.OrdinalIgnoreCase)) {
            operations = Array.Empty<CounterOperation>();
            return true;
        }

        IReadOnlyList<string> tokens = HtmlRenderCssValues.SplitWhitespace(value);
        var parsedOperations = new List<CounterOperation>();
        for (int index = 0; index < tokens.Count; index++) {
            string name = HtmlCssEscapeDecoder.Decode(tokens[index].Trim());
            if (!IsCounterName(name)) {
                operations = Array.Empty<CounterOperation>();
                return false;
            }
            int counterValue = defaultValue;
            if (index + 1 < tokens.Count
                && int.TryParse(tokens[index + 1], NumberStyles.Integer, CultureInfo.InvariantCulture, out int parsed)) {
                counterValue = parsed;
                index++;
            }

            parsedOperations.Add(new CounterOperation(name, counterValue));
        }

        operations = parsedOperations.AsReadOnly();
        return true;
    }

    private static bool TryEvaluate(
        string expression,
        IElement element,
        CounterState counters,
        out string generated,
        out string detail) {
        var text = new StringBuilder();
        int cursor = 0;
        while (cursor < expression.Length) {
            while (cursor < expression.Length && char.IsWhiteSpace(expression[cursor])) cursor++;
            if (cursor >= expression.Length) break;
            char current = expression[cursor];
            if (current == '\'' || current == '"') {
                if (!TryReadQuoted(expression, ref cursor, out string literal)) {
                    generated = string.Empty;
                    detail = "unterminated CSS string";
                    return false;
                }

                text.Append(literal);
                continue;
            }

            if (!TryReadFunction(expression, ref cursor, out string functionName, out string arguments)) {
                generated = string.Empty;
                detail = "unsupported content token near " + expression.Substring(cursor).Trim();
                return false;
            }

            if (string.Equals(functionName, "attr", StringComparison.OrdinalIgnoreCase)) {
                string attributeName = HtmlCssEscapeDecoder.Decode(arguments.Trim());
                if (!IsAttributeName(attributeName)) {
                    generated = string.Empty;
                    detail = "unsupported attr() expression";
                    return false;
                }

                text.Append(element.GetAttribute(attributeName) ?? string.Empty);
            } else if (string.Equals(functionName, "counter", StringComparison.OrdinalIgnoreCase)) {
                IReadOnlyList<string> parts = SplitArguments(arguments);
                if (parts.Count < 1 || parts.Count > 2 || !IsCounterName(parts[0].Trim())) {
                    generated = string.Empty;
                    detail = "unsupported counter() expression";
                    return false;
                }

                string name = HtmlCssEscapeDecoder.Decode(parts[0].Trim());
                string style = parts.Count == 2 ? parts[1].Trim() : "decimal";
                if (!TryFormatCounter(counters.Get(name), style, out string formatted)) {
                    generated = string.Empty;
                    detail = "unsupported counter style " + style;
                    return false;
                }

                text.Append(formatted);
            } else if (string.Equals(functionName, "counters", StringComparison.OrdinalIgnoreCase)) {
                IReadOnlyList<string> parts = SplitArguments(arguments);
                if (!TryParseCountersArguments(parts, out string name, out string separator, out string style)) {
                    generated = string.Empty;
                    detail = "unsupported counters() expression";
                    return false;
                }

                var formattedValues = new List<string>();
                foreach (int value in counters.GetAll(name)) {
                    if (!TryFormatCounter(value, style, out string formatted)) {
                        generated = string.Empty;
                        detail = "unsupported counter style " + style;
                        return false;
                    }

                    formattedValues.Add(formatted);
                }

                text.Append(string.Join(separator, formattedValues));
            } else {
                generated = string.Empty;
                detail = "unsupported generated-content function " + functionName + "()";
                return false;
            }
        }

        generated = text.ToString();
        detail = string.Empty;
        return true;
    }

    private static bool TryReadQuoted(string value, ref int cursor, out string text) {
        char quote = value[cursor++];
        var raw = new StringBuilder();
        while (cursor < value.Length) {
            char current = value[cursor++];
            if (current == quote) {
                text = HtmlCssEscapeDecoder.Decode(raw.ToString());
                return true;
            }

            if (current == '\\' && cursor < value.Length) raw.Append(current).Append(value[cursor++]);
            else raw.Append(current);
        }

        text = string.Empty;
        return false;
    }

    private static bool TryReadFunction(string value, ref int cursor, out string name, out string arguments) {
        int nameStart = cursor;
        while (cursor < value.Length && (char.IsLetterOrDigit(value[cursor]) || value[cursor] == '-' || value[cursor] == '_')) cursor++;
        name = HtmlCssEscapeDecoder.Decode(value.Substring(nameStart, cursor - nameStart));
        if (name.Length == 0 || cursor >= value.Length || value[cursor] != '(') {
            arguments = string.Empty;
            cursor = nameStart;
            return false;
        }

        int open = cursor++;
        int depth = 1;
        char quote = '\0';
        while (cursor < value.Length) {
            char current = value[cursor];
            if (quote != '\0') {
                if (current == quote && !IsEscaped(value, cursor)) quote = '\0';
            } else if (current == '\'' || current == '"') {
                quote = current;
            } else if (current == '(') {
                depth++;
            } else if (current == ')' && --depth == 0) {
                arguments = value.Substring(open + 1, cursor - open - 1);
                cursor++;
                return true;
            }

            cursor++;
        }

        arguments = string.Empty;
        cursor = nameStart;
        return false;
    }

    private static IReadOnlyList<string> SplitArguments(string value) {
        var parts = new List<string>();
        int start = 0;
        int depth = 0;
        char quote = '\0';
        for (int index = 0; index < value.Length; index++) {
            char current = value[index];
            if (quote != '\0') {
                if (current == quote && !IsEscaped(value, index)) quote = '\0';
            } else if (current == '\'' || current == '"') {
                quote = current;
            } else if (current == '(') {
                depth++;
            } else if (current == ')' && depth > 0) {
                depth--;
            } else if (current == ',' && depth == 0) {
                parts.Add(value.Substring(start, index - start).Trim());
                start = index + 1;
            }
        }

        parts.Add(value.Substring(start).Trim());
        return parts.AsReadOnly();
    }

    private static bool TryParseCountersArguments(
        IReadOnlyList<string> parts,
        out string name,
        out string separator,
        out string style) {
        name = string.Empty;
        separator = string.Empty;
        style = "decimal";
        if (parts.Count >= 2 && parts.Count <= 3 && IsCounterName(parts[0].Trim())
            && TryParseCounterSeparator(parts[1], out separator)) {
            name = HtmlCssEscapeDecoder.Decode(parts[0].Trim());
            if (parts.Count == 3) style = parts[2].Trim();
            return true;
        }

        // AngleSharp.Css beta serializes counters(name, "separator", style) as
        // counters(name separator, style), so accept that known normalized form too.
        if (parts.Count < 1 || parts.Count > 2) return false;
        IReadOnlyList<string> normalized = HtmlRenderCssValues.SplitWhitespace(parts[0]);
        if (normalized.Count < 2 || !IsCounterName(normalized[0])) return false;
        name = HtmlCssEscapeDecoder.Decode(normalized[0]);
        string normalizedSeparator = string.Join(" ", normalized.Skip(1));
        if (!TryParseCounterSeparator(normalizedSeparator, out separator)) return false;
        if (parts.Count == 2) style = parts[1].Trim();
        return true;
    }

    private static bool TryParseCounterSeparator(string value, out string separator) {
        string trimmed = value.Trim();
        if (TryParseQuotedValue(trimmed, out separator)) return true;
        if (trimmed.Length == 0 || trimmed.Any(char.IsWhiteSpace)) {
            separator = string.Empty;
            return false;
        }

        separator = HtmlCssEscapeDecoder.Decode(trimmed);
        return true;
    }

    private static bool TryParseQuotedValue(string value, out string result) {
        int cursor = 0;
        string trimmed = value.Trim();
        if (trimmed.Length < 2 || trimmed[0] != '\'' && trimmed[0] != '"') {
            result = string.Empty;
            return false;
        }

        if (!TryReadQuoted(trimmed, ref cursor, out result)) return false;
        while (cursor < trimmed.Length && char.IsWhiteSpace(trimmed[cursor])) cursor++;
        return cursor == trimmed.Length;
    }

    private static bool TryFormatCounter(int value, string style, out string formatted) {
        string normalized = HtmlCssEscapeDecoder.Decode(style.Trim()).ToLowerInvariant();
        switch (normalized) {
            case "decimal-leading-zero":
                formatted = value >= -9 && value <= 9
                    ? value < 0 ? "-0" + (-value).ToString(CultureInfo.InvariantCulture) : "0" + value.ToString(CultureInfo.InvariantCulture)
                    : value.ToString(CultureInfo.InvariantCulture);
                return true;
            case "lower-alpha":
            case "lower-latin":
                formatted = value > 0 ? FormatAlpha(value, false) : value.ToString(CultureInfo.InvariantCulture);
                return true;
            case "upper-alpha":
            case "upper-latin":
                formatted = value > 0 ? FormatAlpha(value, true) : value.ToString(CultureInfo.InvariantCulture);
                return true;
            case "lower-roman":
                formatted = value > 0 && value <= 3999 ? FormatRoman(value).ToLowerInvariant() : value.ToString(CultureInfo.InvariantCulture);
                return true;
            case "upper-roman":
                formatted = value > 0 && value <= 3999 ? FormatRoman(value) : value.ToString(CultureInfo.InvariantCulture);
                return true;
            case "none":
                formatted = string.Empty;
                return true;
            case "decimal":
            case "":
                formatted = value.ToString(CultureInfo.InvariantCulture);
                return true;
            default:
                formatted = string.Empty;
                return false;
        }
    }

    private static string FormatAlpha(int value, bool upper) {
        var result = new StringBuilder();
        int remaining = value;
        char first = upper ? 'A' : 'a';
        while (remaining > 0) {
            remaining--;
            result.Insert(0, (char)(first + remaining % 26));
            remaining /= 26;
        }

        return result.ToString();
    }

    private static string FormatRoman(int value) {
        var result = new StringBuilder();
        int[] values = { 1000, 900, 500, 400, 100, 90, 50, 40, 10, 9, 5, 4, 1 };
        string[] symbols = { "M", "CM", "D", "CD", "C", "XC", "L", "XL", "X", "IX", "V", "IV", "I" };
        for (int index = 0; index < values.Length; index++) {
            while (value >= values[index]) {
                result.Append(symbols[index]);
                value -= values[index];
            }
        }

        return result.ToString();
    }

    private static bool IsCounterName(string value) {
        string name = HtmlCssEscapeDecoder.Decode(value.Trim());
        if (name.Length == 0
            || name == "-"
            || string.Equals(name, "none", StringComparison.OrdinalIgnoreCase)
            || string.Equals(name, "inherit", StringComparison.OrdinalIgnoreCase)
            || string.Equals(name, "initial", StringComparison.OrdinalIgnoreCase)
            || string.Equals(name, "unset", StringComparison.OrdinalIgnoreCase)
            || string.Equals(name, "revert", StringComparison.OrdinalIgnoreCase)
            || string.Equals(name, "revert-layer", StringComparison.OrdinalIgnoreCase)) return false;
        char first = name[0];
        if (!char.IsLetter(first) && first != '_' && first != '-' && first < 0x80) return false;
        return name.All(character => char.IsLetterOrDigit(character) || character == '-' || character == '_');
    }

    private static bool IsAttributeName(string value) =>
        value.Length > 0 && value.All(character => char.IsLetterOrDigit(character) || character == '-' || character == '_' || character == ':');

    private static bool IsEscaped(string value, int index) {
        int slashes = 0;
        for (int current = index - 1; current >= 0 && value[current] == '\\'; current--) slashes++;
        return slashes % 2 == 1;
    }

    private static bool ShouldSkipSubtree(IElement element) {
        string tag = element.TagName.ToLowerInvariant();
        return tag == "head" || tag == "style" || tag == "script" || tag == "template" || tag == "noscript" || tag == "meta" || tag == "link" || tag == "title" || tag == "base";
    }

    private readonly struct CounterOperation {
        internal CounterOperation(string name, int value) {
            Name = name;
            Value = value;
        }

        internal string Name { get; }
        internal int Value { get; }
    }

    private sealed class CounterState {
        private readonly Dictionary<string, List<CounterValue>> _values = new Dictionary<string, List<CounterValue>>(StringComparer.Ordinal);
        private int _nextLevel;

        internal int EnterLevel() => ++_nextLevel;

        internal void ExitLevel(int level) {
            foreach (string name in _values.Keys.ToList()) {
                List<CounterValue> values = _values[name];
                while (values.Count > 0 && values[values.Count - 1].Level == level) values.RemoveAt(values.Count - 1);
                if (values.Count == 0) _values.Remove(name);
            }
        }

        internal void Reset(string name, int value, int level) {
            List<CounterValue> values = GetOrCreate(name);
            if (values.Count > 0 && values[values.Count - 1].Level == level) {
                values[values.Count - 1].Value = value;
            } else {
                values.Add(new CounterValue(level, value));
            }
        }

        internal void Set(string name, int value, int level) {
            List<CounterValue> values = GetOrCreate(name);
            if (values.Count == 0) values.Add(new CounterValue(level, value));
            else values[values.Count - 1].Value = value;
        }

        internal void Increment(string name, int increment, int level) {
            List<CounterValue> values = GetOrCreate(name);
            if (values.Count == 0) values.Add(new CounterValue(level, 0));
            long result = (long)values[values.Count - 1].Value + increment;
            values[values.Count - 1].Value = result > int.MaxValue ? int.MaxValue : result < int.MinValue ? int.MinValue : (int)result;
        }

        internal int Get(string name) {
            return _values.TryGetValue(name, out List<CounterValue>? values) && values.Count > 0
                ? values[values.Count - 1].Value
                : 0;
        }

        internal IReadOnlyList<int> GetAll(string name) {
            return _values.TryGetValue(name, out List<CounterValue>? values)
                ? values.Select(value => value.Value).ToList().AsReadOnly()
                : new[] { 0 };
        }

        private List<CounterValue> GetOrCreate(string name) {
            if (!_values.TryGetValue(name, out List<CounterValue>? values)) {
                values = new List<CounterValue>();
                _values[name] = values;
            }

            return values;
        }
    }

    private sealed class CounterValue {
        internal CounterValue(int level, int value) {
            Level = level;
            Value = value;
        }

        internal int Level { get; }
        internal int Value { get; set; }
    }
}
