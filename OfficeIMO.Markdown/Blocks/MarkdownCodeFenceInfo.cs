namespace OfficeIMO.Markdown;

/// <summary>
/// Parsed fenced-code info string preserving the original raw value while exposing the primary language token.
/// </summary>
public sealed class MarkdownCodeFenceInfo {
    private static readonly IReadOnlyList<string> EmptyClasses = Array.Empty<string>();
    private readonly IReadOnlyDictionary<string, string?> _attributes;
    private readonly IReadOnlyList<string> _classes;

    private MarkdownCodeFenceInfo(
        string infoString,
        string language,
        string additionalInfo,
        string? elementId,
        IReadOnlyList<string> classes,
        IReadOnlyDictionary<string, string?> attributes,
        MarkdownAttributeSet genericAttributes,
        string? genericAttributeSourceText,
        bool hasExplicitAttributes) {
        InfoString = infoString;
        Language = language;
        AdditionalInfo = additionalInfo;
        ElementId = elementId;
        _classes = classes;
        _attributes = attributes;
        GenericAttributes = genericAttributes ?? MarkdownAttributeSet.Empty;
        GenericAttributeSourceText = genericAttributeSourceText;
        HasExplicitAttributes = hasExplicitAttributes;
    }

    /// <summary>
    /// Full normalized fence info string exactly as it should round-trip after the opening fence marker.
    /// </summary>
    public string InfoString { get; }

    /// <summary>
    /// Primary fence language token parsed from the info string.
    /// </summary>
    public string Language { get; }

    /// <summary>
    /// Remaining fence metadata after the primary language token, preserved as raw trimmed text.
    /// </summary>
    public string AdditionalInfo { get; }

    /// <summary>
    /// Optional element id parsed from <c>#id</c> shorthand or <c>id=...</c> metadata.
    /// </summary>
    public string? ElementId { get; }

    /// <summary>
    /// Optional CSS classes parsed from <c>.class</c> shorthand or <c>class=...</c> metadata.
    /// </summary>
    public IReadOnlyList<string> Classes => _classes;

    /// <summary>
    /// Parsed attribute-style metadata from the additional info string.
    /// Recognizes <c>key=value</c>, <c>key="value with spaces"</c>, and standalone flags.
    /// </summary>
    public IReadOnlyDictionary<string, string?> Attributes => _attributes;

    /// <summary>
    /// Markdig-style generic attributes parsed from explicit <c>{...}</c> fence-info attribute blocks only.
    /// Other fence options remain available through <see cref="Attributes"/> for host renderers.
    /// </summary>
    public MarkdownAttributeSet GenericAttributes { get; }

    /// <summary>
    /// Source text for the explicit fence-info attribute block, including braces, when present.
    /// </summary>
    internal string? GenericAttributeSourceText { get; }

    /// <summary>
    /// True when the info string contains explicit attribute syntax such as <c>{...}</c>, <c>#id</c>, or <c>.class</c>.
    /// Plain key/value fence options remain available through <see cref="Attributes"/> for host features but are not projected as generic HTML attributes by ordinary code blocks.
    /// </summary>
    public bool HasExplicitAttributes { get; }

    /// <summary>
    /// Common convenience title resolved from <c>title</c> or <c>caption</c> attributes when present.
    /// </summary>
    public string? Title {
        get {
            if (TryGetAttribute("title", out var title) && !string.IsNullOrWhiteSpace(title)) {
                return title;
            }

            if (TryGetAttribute("caption", out var caption) && !string.IsNullOrWhiteSpace(caption)) {
                return caption;
            }

            return null;
        }
    }

    /// <summary>
    /// Reads a parsed attribute value by key.
    /// </summary>
    public bool TryGetAttribute(string name, out string? value) {
        value = null;
        if (string.IsNullOrWhiteSpace(name) || _attributes.Count == 0) {
            return false;
        }

        return _attributes.TryGetValue(name.Trim(), out value);
    }

    /// <summary>
    /// Reads the first parsed attribute value that matches any of the provided aliases.
    /// </summary>
    public bool TryGetAttribute(out string? value, params string[] aliases) {
        value = null;
        if (aliases == null || aliases.Length == 0 || _attributes.Count == 0) {
            return false;
        }

        for (int i = 0; i < aliases.Length; i++) {
            if (TryGetAttribute(aliases[i], out value)) {
                return true;
            }
        }

        return false;
    }

    /// <summary>
    /// Reads the first parsed attribute value that matches any of the provided aliases.
    /// </summary>
    public string? GetAttribute(params string[] aliases) {
        return TryGetAttribute(out var value, aliases) ? value : null;
    }

    /// <summary>
    /// Attempts to read a parsed boolean attribute value.
    /// </summary>
    public bool TryGetBooleanAttribute(string name, out bool value) {
        value = false;
        return TryGetAttribute(name, out var rawValue) && TryParseBoolean(rawValue, out value);
    }

    /// <summary>
    /// Attempts to read a parsed boolean attribute value from any of the provided aliases.
    /// </summary>
    public bool TryGetBooleanAttribute(out bool value, params string[] aliases) {
        value = false;
        if (aliases == null || aliases.Length == 0 || _attributes.Count == 0) {
            return false;
        }

        for (int i = 0; i < aliases.Length; i++) {
            if (TryGetBooleanAttribute(aliases[i], out value)) {
                return true;
            }
        }

        return false;
    }

    /// <summary>
    /// Attempts to read a parsed 32-bit integer attribute value.
    /// </summary>
    public bool TryGetInt32Attribute(string name, out int value) {
        value = 0;
        return TryGetAttribute(name, out var rawValue)
            && int.TryParse(rawValue, System.Globalization.NumberStyles.Integer, System.Globalization.CultureInfo.InvariantCulture, out value);
    }

    /// <summary>
    /// Attempts to read a parsed 32-bit integer attribute value from any of the provided aliases.
    /// </summary>
    public bool TryGetInt32Attribute(out int value, params string[] aliases) {
        value = 0;
        if (aliases == null || aliases.Length == 0 || _attributes.Count == 0) {
            return false;
        }

        for (int i = 0; i < aliases.Length; i++) {
            if (TryGetInt32Attribute(aliases[i], out value)) {
                return true;
            }
        }

        return false;
    }

    /// <summary>
    /// Determines whether the fence metadata contains the given CSS class.
    /// </summary>
    public bool HasClass(string className) {
        if (string.IsNullOrWhiteSpace(className) || _classes.Count == 0) {
            return false;
        }

        for (int i = 0; i < _classes.Count; i++) {
            if (string.Equals(_classes[i], className.Trim(), StringComparison.OrdinalIgnoreCase)) {
                return true;
            }
        }

        return false;
    }

    /// <summary>
    /// Parses a fenced-code info string into its primary language token and additional metadata.
    /// </summary>
    public static MarkdownCodeFenceInfo Parse(string? infoString) {
        var normalized = (infoString ?? string.Empty).Trim();
        if (normalized.Length == 0) {
            return new MarkdownCodeFenceInfo(
                string.Empty,
                string.Empty,
                string.Empty,
                null,
                EmptyClasses,
                new Dictionary<string, string?>(StringComparer.OrdinalIgnoreCase),
                MarkdownAttributeSet.Empty,
                null,
                hasExplicitAttributes: false);
        }

        if (StartsWithAttributeBlock(normalized)) {
            ParseMetadata(
                normalized,
                out var attributeOnlyElementId,
                out var attributeOnlyClasses,
                out var attributeOnlyAttributes,
                out var attributeOnlyGenericAttributes,
                out var attributeOnlyGenericAttributeSourceText,
                out var attributeOnlyHasExplicitAttributes);
            return new MarkdownCodeFenceInfo(
                normalized,
                string.Empty,
                normalized,
                attributeOnlyElementId,
                attributeOnlyClasses,
                attributeOnlyAttributes,
                attributeOnlyGenericAttributes,
                attributeOnlyGenericAttributeSourceText,
                attributeOnlyHasExplicitAttributes);
        }

        int split = 0;
        while (split < normalized.Length && !char.IsWhiteSpace(normalized[split])) {
            split++;
        }

        var language = DecodeInfoStringToken(normalized.Substring(0, split));
        var additionalInfo = split < normalized.Length
            ? normalized.Substring(split).Trim()
            : string.Empty;

        ParseMetadata(
            additionalInfo,
            out var elementId,
            out var classes,
            out var attributes,
            out var genericAttributes,
            out var genericAttributeSourceText,
            out var hasExplicitAttributes);

        return new MarkdownCodeFenceInfo(
            normalized,
            language,
            additionalInfo,
            elementId,
            classes,
            attributes,
            genericAttributes,
            genericAttributeSourceText,
            hasExplicitAttributes);
    }

    private static bool StartsWithAttributeBlock(string value) {
        if (value.Length == 0 || value[0] != '{') {
            return false;
        }

        int index = 0;
        return TryReadAttributeBlock(value, ref index, out _);
    }

    private static string DecodeInfoStringToken(string token) {
        if (string.IsNullOrEmpty(token)) {
            return string.Empty;
        }

        var decoded = DecodeBackslashEscapes(token);
        return CommonMarkCharacterReference.DecodeAll(decoded);
    }

    private static string DecodeBackslashEscapes(string value) {
        StringBuilder? builder = null;

        for (int i = 0; i < value.Length; i++) {
            char ch = value[i];
            if (ch != '\\' || i + 1 >= value.Length || !IsAsciiPunctuation(value[i + 1])) {
                builder?.Append(ch);
                continue;
            }

            builder ??= new StringBuilder(value.Length).Append(value, 0, i);
            builder.Append(value[i + 1]);
            i++;
        }

        return builder?.ToString() ?? value;
    }

    private static bool IsAsciiPunctuation(char ch) =>
        (ch >= '!' && ch <= '/') ||
        (ch >= ':' && ch <= '@') ||
        (ch >= '[' && ch <= '`') ||
        (ch >= '{' && ch <= '~');

    private static void ParseMetadata(
        string additionalInfo,
        out string? elementId,
        out IReadOnlyList<string> classes,
        out IReadOnlyDictionary<string, string?> attributes,
        out MarkdownAttributeSet genericAttributes,
        out string? genericAttributeSourceText,
        out bool hasExplicitAttributes) {
        elementId = null;
        string? genericElementId = null;
        genericAttributeSourceText = null;
        hasExplicitAttributes = false;
        var parsedAttributes = new Dictionary<string, string?>(StringComparer.OrdinalIgnoreCase);
        var parsedGenericAttributes = new Dictionary<string, string?>(StringComparer.OrdinalIgnoreCase);
        var parsedClasses = new List<string>();
        var parsedGenericClasses = new List<string>();
        int genericAttributeStart = -1;
        int genericAttributeEnd = -1;
        if (string.IsNullOrWhiteSpace(additionalInfo)) {
            classes = EmptyClasses;
            attributes = parsedAttributes;
            genericAttributes = MarkdownAttributeSet.Empty;
            return;
        }

        foreach (var segment in EnumerateMetadataSegments(additionalInfo)) {
            if (segment.IsOpaque) {
                continue;
            }

            if (!segment.IsAttributeBlock) {
                hasExplicitAttributes |= IsExplicitAttributeToken(segment.Value);
                MarkdownGenericAttributeParser.ParseToken(segment.Value, parsedAttributes, parsedClasses, ref elementId);
                continue;
            }

            hasExplicitAttributes = true;
            if (genericAttributeStart < 0) {
                genericAttributeStart = segment.StartIndex;
            }

            genericAttributeEnd = segment.EndIndex;
            MarkdownGenericAttributeParser.ParseTokens(segment.Value, out var blockElementId, out var blockClasses, out var blockAttributes);
            if (string.IsNullOrWhiteSpace(elementId) && !string.IsNullOrWhiteSpace(blockElementId)) {
                elementId = blockElementId;
            }

            if (string.IsNullOrWhiteSpace(genericElementId) && !string.IsNullOrWhiteSpace(blockElementId)) {
                genericElementId = blockElementId;
            }

            for (int i = 0; i < blockClasses.Count; i++) {
                AddClass(parsedClasses, blockClasses[i]);
                AddClass(parsedGenericClasses, blockClasses[i]);
            }

            foreach (var attribute in blockAttributes) {
                if (!parsedAttributes.ContainsKey(attribute.Key)) {
                    parsedAttributes[attribute.Key] = attribute.Value;
                }

                if (!parsedGenericAttributes.ContainsKey(attribute.Key)) {
                    parsedGenericAttributes[attribute.Key] = attribute.Value;
                }
            }
        }

        classes = parsedClasses.Count == 0 ? EmptyClasses : parsedClasses.AsReadOnly();
        attributes = parsedAttributes;
        genericAttributes = MarkdownAttributeSet.Create(genericElementId, parsedGenericClasses, parsedGenericAttributes);
        if (genericAttributeStart >= 0 && genericAttributeEnd >= genericAttributeStart) {
            genericAttributeSourceText = additionalInfo.Substring(genericAttributeStart, genericAttributeEnd - genericAttributeStart + 1);
        }
    }

    private static bool IsExplicitAttributeToken(string token) =>
        !string.IsNullOrWhiteSpace(token) && (token[0] == '#' || token[0] == '.');

    private static IEnumerable<FenceMetadataSegment> EnumerateMetadataSegments(string value) {
        int index = 0;
        while (index < value.Length) {
            while (index < value.Length && char.IsWhiteSpace(value[index])) {
                index++;
            }

            if (index >= value.Length) {
                yield break;
            }

            if (value[index] == '{') {
                var segmentStart = index;
                if (TryReadAttributeBlock(value, ref index, out var block)) {
                    if (!string.IsNullOrWhiteSpace(block)) {
                        yield return new FenceMetadataSegment(block, isAttributeBlock: true, isOpaque: false, segmentStart, index - 1);
                    }

                    continue;
                }

                var opaqueRemainder = value.Substring(index).Trim();
                if (opaqueRemainder.Length > 0) {
                    yield return new FenceMetadataSegment(opaqueRemainder, isAttributeBlock: false, isOpaque: true, index, value.Length - 1);
                }

                yield break;
            }

            int start = index;
            char quote = '\0';
            while (index < value.Length) {
                var ch = value[index];
                if (quote == '\0') {
                    if (char.IsWhiteSpace(ch)) {
                        break;
                    }

                    if (ch == '"' || ch == '\'') {
                        quote = ch;
                    }
                } else if (ch == quote && (index == start || value[index - 1] != '\\')) {
                    quote = '\0';
                }

                index++;
            }

            var token = value.Substring(start, index - start).Trim();
            if (token.Length > 0) {
                yield return new FenceMetadataSegment(token, isAttributeBlock: false, isOpaque: false, start, index - 1);
            }
        }
    }

    private static bool TryReadAttributeBlock(string value, ref int index, out string block) {
        int start = index;
        int depth = 0;
        char quote = '\0';

        while (index < value.Length) {
            var ch = value[index];
            if (quote == '\0') {
                if (ch == '{') {
                    depth++;
                } else if (ch == '}') {
                    depth--;
                    if (depth == 0) {
                        block = value.Substring(start + 1, index - start - 1).Trim();
                        index++;
                        return true;
                    }
                } else if (ch == '"' || ch == '\'') {
                    quote = ch;
                }
            } else if (ch == quote && (index == start || value[index - 1] != '\\')) {
                quote = '\0';
            }

            index++;
        }

        block = string.Empty;
        index = start;
        return false;
    }

    private static void AddClass(IList<string> classes, string? className) {
        if (string.IsNullOrWhiteSpace(className)) {
            return;
        }

        var normalized = className!.Trim();
        for (int i = 0; i < classes.Count; i++) {
            if (string.Equals(classes[i], normalized, StringComparison.OrdinalIgnoreCase)) {
                return;
            }
        }

        classes.Add(normalized);
    }

    private static bool TryParseBoolean(string? rawValue, out bool value) {
        value = false;
        if (string.IsNullOrWhiteSpace(rawValue)) {
            return false;
        }

        switch (rawValue!.Trim().ToLowerInvariant()) {
            case "true":
            case "1":
            case "yes":
            case "on":
                value = true;
                return true;
            case "false":
            case "0":
            case "no":
            case "off":
                value = false;
                return true;
            default:
                return false;
        }
    }

    private readonly struct FenceMetadataSegment {
        public FenceMetadataSegment(string value, bool isAttributeBlock, bool isOpaque, int startIndex, int endIndex) {
            Value = value;
            IsAttributeBlock = isAttributeBlock;
            IsOpaque = isOpaque;
            StartIndex = startIndex;
            EndIndex = endIndex;
        }

        public string Value { get; }

        public bool IsAttributeBlock { get; }

        public bool IsOpaque { get; }

        public int StartIndex { get; }

        public int EndIndex { get; }
    }
}
