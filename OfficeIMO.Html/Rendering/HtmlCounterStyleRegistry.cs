using System.Globalization;
using System.Text;
using AngleSharp.Dom;
using AngleSharp.Html.Dom;

namespace OfficeIMO.Html;

/// <summary>Document-scoped author counter styles shared by lists and generated content.</summary>
internal sealed class HtmlCounterStyleRegistry {
    private const int MaximumFallbackDepth = 32;
    private readonly Dictionary<string, RegisteredDefinition> _definitions = new Dictionary<string, RegisteredDefinition>(StringComparer.Ordinal);

    internal static HtmlCounterStyleRegistry Parse(IHtmlDocument document, HtmlRenderOptions options) {
        var registry = new HtmlCounterStyleRegistry();
        var layers = new CascadeLayerRegistry();
        int sourceOrder = 0;
        foreach (IElement styleElement in document.QuerySelectorAll("style")) {
            string type = (styleElement.GetAttribute("type") ?? string.Empty).Trim();
            int parameter = type.IndexOf(';');
            if (parameter >= 0) type = type.Substring(0, parameter).Trim();
            if (type.Length > 0 && !string.Equals(type, "text/css", StringComparison.OrdinalIgnoreCase)) continue;
            string media = styleElement.GetAttribute("media") ?? string.Empty;
            if (!HtmlComputedStyleEngine.IsApplicableMedia(
                media,
                options.MediaContext,
                options.Mode == HtmlRenderMode.Paged ? options.PageWidth : options.ViewportWidth,
                options.Mode == HtmlRenderMode.Paged ? options.PageHeight : options.ViewportHeight ?? 1056D,
                options.MediaFeatures)) continue;

            string css = styleElement.TextContent;
            IReadOnlyDictionary<int, int> closures = HtmlCssRuleBlockScanner.Scan(css, new HtmlCssProcessingBudget(null));
            registry.Collect(css, 0, css.Length, closures, options, layers, currentLayer: null, ref sourceOrder);
        }
        return registry;
    }

    internal bool TryFormat(int value, string styleName, out string formatted) =>
        TryFormat(value, styleName, out formatted, out _);

    internal bool TryFormat(int value, string styleName, out string formatted, out bool representationLimited) =>
        TryFormat(value, styleName, new HashSet<string>(StringComparer.Ordinal), 0, out formatted, out representationLimited, out _);

    internal bool TryFormatMarker(int value, string styleName, out string marker) =>
        TryFormatMarker(value, styleName, out marker, out _);

    internal bool TryFormatMarker(int value, string styleName, out string marker, out bool representationLimited) {
        marker = string.Empty;
        representationLimited = false;
        string decodedName = HtmlCssEscapeDecoder.Decode(styleName.Trim());
        if (!_definitions.ContainsKey(decodedName)) return false;
        if (!TryFormat(value, decodedName, new HashSet<string>(StringComparer.Ordinal), 0, out string representation, out representationLimited, out string effectiveStyle)) return false;
        marker = _definitions.TryGetValue(effectiveStyle, out RegisteredDefinition? effectiveDefinition)
            ? effectiveDefinition.Value.Prefix + representation + effectiveDefinition.Value.Suffix
            : representation + HtmlCounterStyleFormatter.MarkerSuffix(effectiveStyle);
        if (marker.Length > HtmlCounterStyleFormatter.MaximumGeneratedRepresentationLength) {
            marker = value.ToString(CultureInfo.InvariantCulture) + ". ";
            representationLimited = true;
        }
        return true;
    }

    private bool TryFormat(
        int value,
        string styleName,
        ISet<string> active,
        int depth,
        out string formatted,
        out bool representationLimited,
        out string effectiveStyle) {
        formatted = string.Empty;
        representationLimited = false;
        effectiveStyle = string.Empty;
        string decodedName = HtmlCssEscapeDecoder.Decode(styleName.Trim());
        if (!_definitions.TryGetValue(decodedName, out RegisteredDefinition? registered)) return false;
        Definition definition = registered.Value;
        if (depth >= MaximumFallbackDepth || !active.Add(decodedName)) {
            effectiveStyle = "decimal";
            return HtmlCounterStyleFormatter.TryFormat(value, effectiveStyle, out formatted, out representationLimited);
        }
        try {
            if (definition.IsInRange(value) && definition.TryFormat(value, out formatted, out representationLimited)) {
                effectiveStyle = decodedName;
                return true;
            }
            if (representationLimited) {
                effectiveStyle = "decimal";
                HtmlCounterStyleFormatter.TryFormat(value, "decimal", out formatted);
                return true;
            }
            string fallback = definition.Fallback.Length == 0 ? "decimal" : definition.Fallback;
            if (_definitions.ContainsKey(fallback)) return TryFormat(value, fallback, active, depth + 1, out formatted, out representationLimited, out effectiveStyle);
            if (HtmlCounterStyleFormatter.TryFormat(value, fallback, out formatted, out representationLimited)) {
                effectiveStyle = fallback;
                return true;
            }
            effectiveStyle = "decimal";
            return HtmlCounterStyleFormatter.TryFormat(value, effectiveStyle, out formatted);
        } finally {
            active.Remove(decodedName);
        }
    }

    private void Collect(
        string css,
        int start,
        int end,
        IReadOnlyDictionary<int, int> closures,
        HtmlRenderOptions options,
        CascadeLayerRegistry layers,
        string? currentLayer,
        ref int sourceOrder) {
        int cursor = start;
        while (cursor < end) {
            SkipTrivia(css, ref cursor, end);
            if (cursor >= end) break;
            int ruleStart = cursor;
            bool atRule = css[cursor] == '@';
            if (atRule) cursor++;
            int nameStart = cursor;
            while (cursor < end && IsIdentifierCharacter(css[cursor])) cursor++;
            string ruleName = atRule ? css.Substring(nameStart, cursor - nameStart).ToLowerInvariant() : string.Empty;
            int delimiter = FindRuleDelimiter(css, cursor, end, out char delimiterCharacter);
            if (delimiter < 0) break;
            string prelude = css.Substring(cursor, delimiter - cursor).Trim();
            if (delimiterCharacter == ';') {
                if (ruleName == "layer") layers.RegisterStatement(prelude, currentLayer);
                cursor = delimiter + 1;
                continue;
            }
            if (!closures.TryGetValue(delimiter, out int close) || close >= end) break;

            if (ruleName == "counter-style") {
                Definition? definition = Definition.TryCreate(prelude, css.Substring(delimiter + 1, close - delimiter - 1));
                if (definition != null) RegisterDefinition(
                    definition,
                    currentLayer == null ? null : layers.GetOrder(currentLayer),
                    sourceOrder++);
            } else if (ruleName == "media") {
                if (HtmlComputedStyleEngine.IsApplicableMedia(
                    prelude,
                    options.MediaContext,
                    options.Mode == HtmlRenderMode.Paged ? options.PageWidth : options.ViewportWidth,
                    options.Mode == HtmlRenderMode.Paged ? options.PageHeight : options.ViewportHeight ?? 1056D,
                    options.MediaFeatures)) Collect(css, delimiter + 1, close, closures, options, layers, currentLayer, ref sourceOrder);
            } else if (ruleName == "supports") {
                if (HtmlComputedStyleEngine.IsApplicableSupports(prelude)) Collect(css, delimiter + 1, close, closures, options, layers, currentLayer, ref sourceOrder);
            } else if (ruleName == "layer") {
                (string nestedLayer, _) = layers.RegisterBlock(prelude, currentLayer);
                Collect(css, delimiter + 1, close, closures, options, layers, nestedLayer, ref sourceOrder);
            }

            cursor = close + 1;
            if (cursor <= ruleStart) cursor = ruleStart + 1;
        }
    }

    private void RegisterDefinition(Definition definition, CascadeLayerOrder? layerOrder, int sourceOrder) {
        var candidate = new RegisteredDefinition(definition, layerOrder, sourceOrder);
        if (!_definitions.TryGetValue(definition.Name, out RegisteredDefinition? existing)
            || ShouldReplace(existing, candidate)) {
            _definitions[definition.Name] = candidate;
        }
    }

    private static bool ShouldReplace(RegisteredDefinition existing, RegisteredDefinition candidate) {
        if ((existing.LayerOrder != null) != (candidate.LayerOrder != null)) return candidate.LayerOrder == null;
        if (existing.LayerOrder != null && candidate.LayerOrder != null) {
            int layerComparison = candidate.LayerOrder.CompareTo(existing.LayerOrder);
            if (layerComparison != 0) return layerComparison > 0;
        }
        return candidate.SourceOrder >= existing.SourceOrder;
    }

    private static int FindRuleDelimiter(string css, int start, int end, out char delimiter) {
        delimiter = '\0';
        char quote = '\0';
        int parenthesisDepth = 0;
        for (int index = start; index < end; index++) {
            char current = css[index];
            if (current == '/' && index + 1 < end && css[index + 1] == '*') {
                int commentEnd = css.IndexOf("*/", index + 2, StringComparison.Ordinal);
                if (commentEnd < 0 || commentEnd >= end) return -1;
                index = commentEnd + 1;
                continue;
            }
            if (current == '\\') { index++; continue; }
            if (quote != '\0') { if (current == quote) quote = '\0'; continue; }
            if (current is '\'' or '"') quote = current;
            else if (current == '(') parenthesisDepth++;
            else if (current == ')' && parenthesisDepth > 0) parenthesisDepth--;
            else if (parenthesisDepth == 0 && (current == '{' || current == ';')) {
                delimiter = current;
                return index;
            }
        }
        return -1;
    }

    private static void SkipTrivia(string css, ref int cursor, int end) {
        while (cursor < end) {
            if (char.IsWhiteSpace(css[cursor])) { cursor++; continue; }
            if (css[cursor] == '/' && cursor + 1 < end && css[cursor + 1] == '*') {
                int commentEnd = css.IndexOf("*/", cursor + 2, StringComparison.Ordinal);
                cursor = commentEnd < 0 || commentEnd >= end ? end : commentEnd + 2;
                continue;
            }
            break;
        }
    }

    private static bool IsIdentifierCharacter(char value) =>
        char.IsLetterOrDigit(value) || value == '-' || value == '_' || value >= 0x80;

    private sealed class RegisteredDefinition {
        internal RegisteredDefinition(Definition value, CascadeLayerOrder? layerOrder, int sourceOrder) {
            Value = value;
            LayerOrder = layerOrder;
            SourceOrder = sourceOrder;
        }

        internal Definition Value { get; }
        internal CascadeLayerOrder? LayerOrder { get; }
        internal int SourceOrder { get; }
    }

    private sealed class Definition {
        private readonly string _system;
        private readonly int _fixedFirst;
        private readonly IReadOnlyList<string> _symbols;
        private readonly IReadOnlyList<AdditiveSymbol> _additiveSymbols;
        private readonly IReadOnlyList<ValueRange> _ranges;
        private readonly string _negativePrefix;
        private readonly string _negativeSuffix;
        private readonly int _padWidth;
        private readonly string _padSymbol;

        private Definition(
            string name,
            string system,
            int fixedFirst,
            IReadOnlyList<string> symbols,
            IReadOnlyList<AdditiveSymbol> additiveSymbols,
            IReadOnlyList<ValueRange> ranges,
            string negativePrefix,
            string negativeSuffix,
            int padWidth,
            string padSymbol,
            string prefix,
            string suffix,
            string fallback) {
            Name = name;
            _system = system;
            _fixedFirst = fixedFirst;
            _symbols = symbols;
            _additiveSymbols = additiveSymbols;
            _ranges = ranges;
            _negativePrefix = negativePrefix;
            _negativeSuffix = negativeSuffix;
            _padWidth = padWidth;
            _padSymbol = padSymbol;
            Prefix = prefix;
            Suffix = suffix;
            Fallback = fallback;
        }

        internal string Name { get; }
        internal string Prefix { get; }
        internal string Suffix { get; }
        internal string Fallback { get; }

        internal static Definition? TryCreate(string rawName, string body) {
            string name = HtmlCssEscapeDecoder.Decode(rawName.Trim());
            IReadOnlyDictionary<string, string> descriptors = ParseDescriptors(body);
            if (name.Length == 0) return null;
            IReadOnlyList<string> systemParts = HtmlRenderCssValues.SplitWhitespace(GetDescriptor(descriptors, "system"));
            if (!TryParseSystemDescriptor(systemParts, out string system, out int fixedFirst)) {
                system = "symbolic";
                fixedFirst = 1;
            }

            if (!TryParseSymbols(GetDescriptor(descriptors, "symbols"), out IReadOnlyList<string> symbols)) {
                symbols = Array.Empty<string>();
            }
            if (!TryParseAdditiveSymbols(GetDescriptor(descriptors, "additive-symbols"), out IReadOnlyList<AdditiveSymbol> additive)) {
                additive = Array.Empty<AdditiveSymbol>();
            }
            if (system == "additive" && additive.Count == 0) return null;
            if (system != "additive" && symbols.Count == 0) return null;
            if (system is "numeric" or "alphabetic" && symbols.Count < 2) return null;
            if (!TryParseRanges(GetDescriptor(descriptors, "range"), out IReadOnlyList<ValueRange> ranges)) {
                ranges = Array.Empty<ValueRange>();
            }
            if (!TryParseStringPair(GetDescriptor(descriptors, "negative"), "-", string.Empty, out string negativePrefix, out string negativeSuffix)) {
                negativePrefix = "-";
                negativeSuffix = string.Empty;
            }
            if (!TryParsePad(GetDescriptor(descriptors, "pad"), out int padWidth, out string padSymbol)) {
                padWidth = 0;
                padSymbol = string.Empty;
            }
            if (!TryParseSingleString(GetDescriptor(descriptors, "prefix"), string.Empty, out string prefix)) {
                prefix = string.Empty;
            }
            if (!TryParseSingleString(GetDescriptor(descriptors, "suffix"), ". ", out string suffix)) {
                suffix = ". ";
            }
            string fallback = HtmlCssEscapeDecoder.Decode(GetDescriptor(descriptors, "fallback").Trim());
            return new Definition(name, system, fixedFirst, symbols, additive, ranges, negativePrefix, negativeSuffix, padWidth, padSymbol, prefix, suffix, fallback);
        }

        private static bool TryParseSystemDescriptor(IReadOnlyList<string> parts, out string system, out int fixedFirst) {
            system = "symbolic";
            fixedFirst = 1;
            if (parts.Count == 0) return true;

            string requested = parts[0].ToLowerInvariant();
            if (requested == "fixed") {
                if (parts.Count == 1) {
                    system = requested;
                } else if (parts.Count == 2
                    && int.TryParse(parts[1], NumberStyles.Integer, CultureInfo.InvariantCulture, out int parsedFirst)) {
                    system = requested;
                    fixedFirst = parsedFirst;
                }
                return true;
            }

            if (requested is not ("cyclic" or "numeric" or "alphabetic" or "symbolic" or "additive")) return false;
            if (parts.Count == 1) system = requested;
            return true;
        }

        private static IReadOnlyDictionary<string, string> ParseDescriptors(string body) {
            var parsed = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            body = HtmlComputedStyleEngine.StripCssCommentsOutsideStrings(body);
            foreach (string declaration in HtmlRenderCssValues.SplitTopLevel(body, ';')) {
                int colon = FindTopLevelColon(declaration);
                if (colon <= 0) continue;
                string property = declaration.Substring(0, colon).Trim();
                string value = declaration.Substring(colon + 1).Trim();
                if (property.Length > 0 && value.Length > 0 && IsValidDescriptor(property, value)) parsed[property] = value;
            }
            return parsed;
        }

        private static bool IsValidDescriptor(string property, string value) {
            switch (property.ToLowerInvariant()) {
                case "system":
                    return TryParseSystemDescriptor(HtmlRenderCssValues.SplitWhitespace(value), out _, out _);
                case "symbols":
                    return TryParseSymbols(value, out IReadOnlyList<string> symbols) && symbols.Count > 0;
                case "additive-symbols":
                    return TryParseAdditiveSymbols(value, out IReadOnlyList<AdditiveSymbol> additive) && additive.Count > 0;
                case "range":
                    return TryParseRanges(value, out _);
                case "negative":
                    return TryParseStringPair(value, "-", string.Empty, out _, out _);
                case "pad":
                    return TryParsePad(value, out _, out _);
                case "prefix":
                case "suffix":
                    return TryParseSingleString(value, string.Empty, out _);
                case "fallback":
                    return HtmlCssEscapeDecoder.Decode(value.Trim()).Length > 0;
                default:
                    return false;
            }
        }

        private static int FindTopLevelColon(string value) {
            char quote = '\0';
            int depth = 0;
            for (int index = 0; index < value.Length; index++) {
                char current = value[index];
                if (current == '\\') { index++; continue; }
                if (quote != '\0') { if (current == quote) quote = '\0'; continue; }
                if (current is '\'' or '"') quote = current;
                else if (current == '(') depth++;
                else if (current == ')' && depth > 0) depth--;
                else if (current == ':' && depth == 0) return index;
            }
            return -1;
        }

        private static string GetDescriptor(IReadOnlyDictionary<string, string> descriptors, string name) =>
            descriptors.TryGetValue(name, out string? value) ? value : string.Empty;

        internal bool IsInRange(int value) => _ranges.Count > 0
            ? _ranges.Any(range => range.Contains(value))
            : _system != "additive" || value >= 0;

        internal bool TryFormat(int value, out string formatted, out bool representationLimited) {
            formatted = string.Empty;
            representationLimited = false;
            bool negative = value < 0;
            long magnitude = Math.Abs((long)value);
            string representation;
            switch (_system) {
                case "cyclic":
                    int cyclic = (int)(((value - 1L) % _symbols.Count + _symbols.Count) % _symbols.Count);
                    representation = _symbols[cyclic];
                    negative = false;
                    break;
                case "fixed":
                    long fixedIndex = (long)value - _fixedFirst;
                    if (fixedIndex < 0 || fixedIndex >= _symbols.Count) return false;
                    representation = _symbols[(int)fixedIndex];
                    negative = false;
                    break;
                case "numeric":
                    representation = FormatNumericMagnitude(magnitude, _symbols);
                    break;
                case "alphabetic":
                    if (value <= 0) return false;
                    representation = HtmlCounterStyleFormatter.FormatAlphabetic(value, _symbols);
                    negative = false;
                    break;
                case "symbolic":
                    if (value <= 0) return false;
                    int symbolIndex = (value - 1) % _symbols.Count;
                    int repetitions = ((value - 1) / _symbols.Count) + 1;
                    if (!HtmlCounterStyleFormatter.TryRepeatSymbol(_symbols[symbolIndex], repetitions, out representation)) {
                        representationLimited = true;
                        return false;
                    }
                    negative = false;
                    break;
                default:
                    if (!TryFormatAdditive(magnitude, out representation, out representationLimited)) return false;
                    break;
            }

            if (representation.Length > HtmlCounterStyleFormatter.MaximumGeneratedRepresentationLength) {
                representationLimited = true;
                return false;
            }
            int symbolCount = CountTextElements(representation);
            if (negative) {
                symbolCount += CountTextElements(_negativePrefix) + CountTextElements(_negativeSuffix);
            }
            if (_padWidth > symbolCount) {
                if (!HtmlCounterStyleFormatter.TryRepeatSymbol(_padSymbol, _padWidth - symbolCount, out string padding)
                    || padding.Length + representation.Length > HtmlCounterStyleFormatter.MaximumGeneratedRepresentationLength) {
                    representationLimited = true;
                    return false;
                }
                representation = padding + representation;
            }
            formatted = negative ? _negativePrefix + representation + _negativeSuffix : representation;
            if (formatted.Length > HtmlCounterStyleFormatter.MaximumGeneratedRepresentationLength) {
                formatted = string.Empty;
                representationLimited = true;
                return false;
            }
            return true;
        }

        private bool TryFormatAdditive(long value, out string representation, out bool representationLimited) {
            representationLimited = false;
            var result = new StringBuilder();
            long remaining = value;
            foreach (AdditiveSymbol additive in _additiveSymbols) {
                if (additive.Weight == 0) {
                    if (remaining == 0 && result.Length == 0) result.Append(additive.Symbol);
                    continue;
                }
                long count = remaining / additive.Weight;
                if (count <= 0) continue;
                if (count > HtmlCounterStyleFormatter.MaximumGeneratedRepresentationLength
                    || additive.Symbol.Length > 0
                    && count > (HtmlCounterStyleFormatter.MaximumGeneratedRepresentationLength - result.Length) / additive.Symbol.Length) {
                    representation = string.Empty;
                    representationLimited = true;
                    return false;
                }
                for (long index = 0; index < count; index++) result.Append(additive.Symbol);
                remaining %= additive.Weight;
            }
            representation = result.ToString();
            return remaining == 0 && representation.Length > 0;
        }

        private static string FormatNumericMagnitude(long value, IReadOnlyList<string> symbols) {
            if (value == 0) return symbols[0];
            var parts = new List<string>();
            long remaining = value;
            while (remaining > 0) {
                parts.Add(symbols[(int)(remaining % symbols.Count)]);
                remaining /= symbols.Count;
            }
            parts.Reverse();
            return string.Concat(parts);
        }

        private static int CountTextElements(string value) => StringInfo.ParseCombiningCharacters(value).Length;

        private static bool TryParseSymbols(string? value, out IReadOnlyList<string> symbols) {
            symbols = Array.Empty<string>();
            if (string.IsNullOrWhiteSpace(value)) return true;
            if (!HtmlCounterStyleFormatter.TryTokenizeSymbols(value!, out IReadOnlyList<string> tokens)) return false;
            var parsed = new List<string>();
            foreach (string token in tokens) {
                if (!TryParseSymbol(token, out string symbol)) return false;
                parsed.Add(symbol);
            }
            symbols = parsed.AsReadOnly();
            return true;
        }

        private static bool TryParseAdditiveSymbols(string? value, out IReadOnlyList<AdditiveSymbol> symbols) {
            symbols = Array.Empty<AdditiveSymbol>();
            if (string.IsNullOrWhiteSpace(value)) return true;
            var parsed = new List<AdditiveSymbol>();
            long previous = long.MaxValue;
            foreach (string entry in HtmlRenderCssValues.SplitTopLevelCommas(value)) {
                IReadOnlyList<string> parts = HtmlRenderCssValues.SplitWhitespace(entry);
                if (parts.Count != 2
                    || !long.TryParse(parts[0], NumberStyles.None, CultureInfo.InvariantCulture, out long weight)
                    || weight < 0 || weight >= previous
                    || !TryParseSymbol(parts[1], out string symbol)) return false;
                parsed.Add(new AdditiveSymbol(weight, symbol));
                previous = weight;
            }
            symbols = parsed.AsReadOnly();
            return true;
        }

        private static bool TryParseRanges(string? value, out IReadOnlyList<ValueRange> ranges) {
            ranges = Array.Empty<ValueRange>();
            if (string.IsNullOrWhiteSpace(value) || string.Equals(value!.Trim(), "auto", StringComparison.OrdinalIgnoreCase)) return true;
            var parsed = new List<ValueRange>();
            foreach (string entry in HtmlRenderCssValues.SplitTopLevelCommas(value)) {
                IReadOnlyList<string> parts = HtmlRenderCssValues.SplitWhitespace(entry);
                if (parts.Count != 2 || !TryParseRangeBound(parts[0], out long minimum) || !TryParseRangeBound(parts[1], out long maximum) || minimum > maximum) return false;
                parsed.Add(new ValueRange(minimum, maximum));
            }
            ranges = parsed.AsReadOnly();
            return true;
        }

        private static bool TryParseRangeBound(string value, out long parsed) {
            if (string.Equals(value, "infinite", StringComparison.OrdinalIgnoreCase)) { parsed = long.MaxValue; return true; }
            if (string.Equals(value, "-infinite", StringComparison.OrdinalIgnoreCase)) { parsed = long.MinValue; return true; }
            return long.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out parsed);
        }

        private static bool TryParseStringPair(string? value, string defaultFirst, string defaultSecond, out string first, out string second) {
            first = defaultFirst;
            second = defaultSecond;
            if (string.IsNullOrWhiteSpace(value)) return true;
            IReadOnlyList<string> parts = HtmlRenderCssValues.SplitWhitespace(value);
            if (parts.Count is < 1 or > 2 || !TryParseAffix(parts[0], out first)) return false;
            if (parts.Count == 2 && !TryParseAffix(parts[1], out second)) return false;
            return true;
        }

        private static bool TryParsePad(string? value, out int width, out string symbol) {
            width = 0;
            symbol = string.Empty;
            if (string.IsNullOrWhiteSpace(value)) return true;
            IReadOnlyList<string> parts = HtmlRenderCssValues.SplitWhitespace(value);
            return parts.Count == 2
                && int.TryParse(parts[0], NumberStyles.None, CultureInfo.InvariantCulture, out width)
                && width >= 0 && width <= 4096
                && TryParseSymbol(parts[1], out symbol);
        }

        private static bool TryParseSingleString(string? value, string defaultValue, out string parsed) {
            parsed = defaultValue;
            return string.IsNullOrWhiteSpace(value) || TryParseAffix(value!.Trim(), out parsed);
        }

        private static bool TryParseAffix(string value, out string affix) {
            if (HtmlCounterStyleFormatter.TryUnquote(value, out affix)) {
                affix = HtmlCssEscapeDecoder.Decode(affix);
                return true;
            }

            affix = HtmlCssEscapeDecoder.Decode(value.Trim());
            return affix.Length > 0 && affix.IndexOfAny(new[] { ' ', '\t', '\r', '\n' }) < 0;
        }

        private static bool TryParseSymbol(string value, out string symbol) {
            if (HtmlCounterStyleFormatter.TryUnquote(value, out symbol)) {
                symbol = HtmlCssEscapeDecoder.Decode(symbol);
                return symbol.Length > 0;
            }
            symbol = HtmlCssEscapeDecoder.Decode(value.Trim());
            return symbol.Length > 0 && symbol.IndexOfAny(new[] { ' ', '\t', '\r', '\n' }) < 0;
        }
    }

    private readonly struct AdditiveSymbol {
        internal AdditiveSymbol(long weight, string symbol) { Weight = weight; Symbol = symbol; }
        internal long Weight { get; }
        internal string Symbol { get; }
    }

    private readonly struct ValueRange {
        internal ValueRange(long minimum, long maximum) { Minimum = minimum; Maximum = maximum; }
        internal long Minimum { get; }
        internal long Maximum { get; }
        internal bool Contains(int value) => value >= Minimum && value <= Maximum;
    }
}
