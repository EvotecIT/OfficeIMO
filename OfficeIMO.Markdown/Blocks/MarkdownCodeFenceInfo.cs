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
        IReadOnlyDictionary<string, string?> attributes) {
        InfoString = infoString;
        Language = language;
        AdditionalInfo = additionalInfo;
        ElementId = elementId;
        _classes = classes;
        _attributes = attributes;
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
                new Dictionary<string, string?>(StringComparer.OrdinalIgnoreCase));
        }

        int split = 0;
        while (split < normalized.Length && !char.IsWhiteSpace(normalized[split])) {
            split++;
        }

        var language = normalized.Substring(0, split);
        var additionalInfo = split < normalized.Length
            ? normalized.Substring(split).Trim()
            : string.Empty;

        ParseMetadata(additionalInfo, out var elementId, out var classes, out var attributes);

        return new MarkdownCodeFenceInfo(normalized, language, additionalInfo, elementId, classes, attributes);
    }

    private static void ParseMetadata(
        string additionalInfo,
        out string? elementId,
        out IReadOnlyList<string> classes,
        out IReadOnlyDictionary<string, string?> attributes) {
        elementId = null;
        var parsedAttributes = new Dictionary<string, string?>(StringComparer.OrdinalIgnoreCase);
        var parsedClasses = new List<string>();
        if (string.IsNullOrWhiteSpace(additionalInfo)) {
            classes = EmptyClasses;
            attributes = parsedAttributes;
            return;
        }

        foreach (var segment in EnumerateMetadataSegments(additionalInfo)) {
            if (segment.IsOpaque) {
                continue;
            }

            if (!segment.IsAttributeBlock) {
                ParseMetadataToken(segment.Value, parsedAttributes, parsedClasses, ref elementId);
                continue;
            }

            foreach (var token in Tokenize(segment.Value)) {
                ParseMetadataToken(token, parsedAttributes, parsedClasses, ref elementId);
            }
        }

        classes = parsedClasses.Count == 0 ? EmptyClasses : parsedClasses.AsReadOnly();
        attributes = parsedAttributes;
    }

    private static IEnumerable<string> Tokenize(string value) {
        var current = new StringBuilder();
        char quote = '\0';

        for (int i = 0; i < value.Length; i++) {
            var ch = value[i];
            if (quote == '\0') {
                if (char.IsWhiteSpace(ch)) {
                    if (current.Length > 0) {
                        yield return current.ToString();
                        current.Clear();
                    }

                    continue;
                }

                if (ch == '"' || ch == '\'') {
                    quote = ch;
                }

                current.Append(ch);
                continue;
            }

            current.Append(ch);
            if (ch == quote && (i == 0 || value[i - 1] != '\\')) {
                quote = '\0';
            }
        }

        if (current.Length > 0) {
            yield return current.ToString();
        }
    }

    private static string? Unquote(string? value) {
        if (string.IsNullOrEmpty(value)) {
            return value;
        }

        if (value!.Length >= 2) {
            var first = value[0];
            var last = value[value.Length - 1];
            if ((first == '"' && last == '"') || (first == '\'' && last == '\'')) {
                return value.Substring(1, value.Length - 2);
            }
        }

        return value;
    }

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
                if (TryReadAttributeBlock(value, ref index, out var block)) {
                    if (!string.IsNullOrWhiteSpace(block)) {
                        yield return new FenceMetadataSegment(block, isAttributeBlock: true, isOpaque: false);
                    }

                    continue;
                }

                var opaqueRemainder = value.Substring(index).Trim();
                if (opaqueRemainder.Length > 0) {
                    yield return new FenceMetadataSegment(opaqueRemainder, isAttributeBlock: false, isOpaque: true);
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
                yield return new FenceMetadataSegment(token, isAttributeBlock: false, isOpaque: false);
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

    private static void ParseMetadataToken(
        string token,
        IDictionary<string, string?> attributes,
        IList<string> classes,
        ref string? elementId) {
        var trimmed = token.Trim();
        if (string.IsNullOrWhiteSpace(trimmed)) {
            return;
        }

        if (trimmed[0] == '#' && trimmed.Length > 1) {
            if (string.IsNullOrWhiteSpace(elementId)) {
                elementId = trimmed.Substring(1);
            }

            return;
        }

        if (trimmed[0] == '.' && trimmed.Length > 1) {
            AddClass(classes, trimmed.Substring(1));
            return;
        }

        int equals = trimmed.IndexOf('=');
        if (equals > 0) {
            var key = trimmed.Substring(0, equals).Trim();
            if (key.Length == 0 || attributes.ContainsKey(key)) {
                return;
            }

            var rawValue = trimmed.Substring(equals + 1).Trim();
            var value = Unquote(rawValue);
            attributes[key] = value;

            if (string.Equals(key, "id", StringComparison.OrdinalIgnoreCase) && string.IsNullOrWhiteSpace(elementId) && !string.IsNullOrWhiteSpace(value)) {
                elementId = value;
                return;
            }

            if (string.Equals(key, "class", StringComparison.OrdinalIgnoreCase) && !string.IsNullOrWhiteSpace(value)) {
                var classList = value!;
                foreach (var className in classList.Split(new[] { ' ', '\t', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries)) {
                    AddClass(classes, className);
                }
            }

            return;
        }

        if (!attributes.ContainsKey(trimmed)) {
            attributes[trimmed] = "true";
        }
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
        public FenceMetadataSegment(string value, bool isAttributeBlock, bool isOpaque) {
            Value = value;
            IsAttributeBlock = isAttributeBlock;
            IsOpaque = isOpaque;
        }

        public string Value { get; }

        public bool IsAttributeBlock { get; }

        public bool IsOpaque { get; }
    }
}
