namespace OfficeIMO.AsciiDoc;

internal static class AsciiDocLineClassifier {
    internal static bool IsBlank(string content) {
        for (int index = 0; index < content.Length; index++) {
            if (!char.IsWhiteSpace(content[index])) return false;
        }
        return true;
    }

    internal static bool IsLineComment(string content) =>
        content.Length >= 2 && content[0] == '/' && content[1] == '/';

    internal static bool TryGetDelimiter(string content, out AsciiDocDelimitedBlockKind kind) {
        if (IsRepeatedDelimiter(content, '-', 4)) { kind = AsciiDocDelimitedBlockKind.Listing; return true; }
        if (IsRepeatedDelimiter(content, '.', 4)) { kind = AsciiDocDelimitedBlockKind.Literal; return true; }
        if (IsRepeatedDelimiter(content, '=', 4)) { kind = AsciiDocDelimitedBlockKind.Example; return true; }
        if (IsRepeatedDelimiter(content, '*', 4)) { kind = AsciiDocDelimitedBlockKind.Sidebar; return true; }
        if (IsRepeatedDelimiter(content, '_', 4)) { kind = AsciiDocDelimitedBlockKind.Quote; return true; }
        if (IsRepeatedDelimiter(content, '+', 4)) { kind = AsciiDocDelimitedBlockKind.Passthrough; return true; }
        if (IsRepeatedDelimiter(content, '/', 4)) { kind = AsciiDocDelimitedBlockKind.Comment; return true; }
        switch (content) {
            case "--": kind = AsciiDocDelimitedBlockKind.Open; return true;
            case "|===": kind = AsciiDocDelimitedBlockKind.Table; return true;
            case ",===": kind = AsciiDocDelimitedBlockKind.Table; return true;
            case ":===": kind = AsciiDocDelimitedBlockKind.Table; return true;
            default:
                kind = default;
                return false;
        }
    }

    private static bool IsRepeatedDelimiter(string content, char marker, int minimumLength) {
        if (content.Length < minimumLength) return false;
        for (int index = 0; index < content.Length; index++) {
            if (content[index] != marker) return false;
        }
        return true;
    }

    internal static bool TryParseHeading(string content, out int markerLength, out int titleStart) {
        markerLength = 0;
        titleStart = 0;
        while (markerLength < content.Length && markerLength < 6 && content[markerLength] == '=') markerLength++;
        if (markerLength == 0 || markerLength >= content.Length || content[markerLength] != ' ') return false;
        titleStart = markerLength + 1;
        return titleStart <= content.Length;
    }

    internal static bool TryParseAttribute(string content, out AttributeParts parts) {
        parts = default;
        if (content.Length < 3 || content[0] != ':') return false;
        int separator = content.IndexOf(':', 1);
        if (separator < 2) return false;

        int rawNameStart = 1;
        int rawNameLength = separator - 1;
        bool isUnset = false;
        if (rawNameLength > 0 && content[rawNameStart] == '!') {
            isUnset = true;
            rawNameStart++;
            rawNameLength--;
        }
        if (rawNameLength > 0 && content[rawNameStart + rawNameLength - 1] == '!') {
            isUnset = true;
            rawNameLength--;
        }
        if (rawNameLength == 0) return false;

        string name = content.Substring(rawNameStart, rawNameLength);
        if (!AsciiDocText.IsAttributeName(name)) return false;

        int valueStart = separator + 1;
        while (valueStart < content.Length && (content[valueStart] == ' ' || content[valueStart] == '\t')) valueStart++;
        parts = new AttributeParts(name, content.Substring(valueStart), isUnset, rawNameStart, rawNameLength, separator, valueStart);
        return true;
    }

    internal static bool TryParseBlockAttributeList(string content, out string value) {
        value = string.Empty;
        if (content.Length < 2 || content[0] != '[' || content[content.Length - 1] != ']') return false;
        if (content.Length >= 4 && content[1] == '[' && content[content.Length - 2] == ']') return false;
        value = content.Substring(1, content.Length - 2);
        return true;
    }

    internal static bool TryParseBlockTitle(string content, out string title) {
        title = string.Empty;
        if (content.Length < 2 || content[0] != '.' || char.IsWhiteSpace(content[1]) || content[1] == '.') return false;
        title = content.Substring(1);
        return true;
    }

    internal static bool TryParseBlockAnchor(string content, out string id, out string? referenceText) {
        id = string.Empty;
        referenceText = null;
        if (content.Length < 5 || !content.StartsWith("[[", StringComparison.Ordinal) || !content.EndsWith("]]", StringComparison.Ordinal)) return false;
        string value = content.Substring(2, content.Length - 4);
        int comma = value.IndexOf(',');
        id = comma < 0 ? value : value.Substring(0, comma);
        referenceText = comma < 0 ? null : value.Substring(comma + 1);
        return id.Length > 0;
    }

    internal static bool TryParseAdmonition(string content, out AdmonitionParts parts) {
        parts = default;
        string[] labels = { "NOTE", "TIP", "IMPORTANT", "WARNING", "CAUTION" };
        for (int index = 0; index < labels.Length; index++) {
            string label = labels[index];
            if (!content.StartsWith(label + ":", StringComparison.Ordinal)) continue;
            int textStart = label.Length + 1;
            if (textStart < content.Length && content[textStart] == ' ') textStart++;
            parts = new AdmonitionParts((AsciiDocAdmonitionKind)index, label, textStart, content.Substring(textStart));
            return true;
        }
        return false;
    }

    internal static bool TryParseDescriptionListItem(string content, out DescriptionListParts parts) {
        parts = default;
        for (int index = 1; index < content.Length - 1; index++) {
            if (content[index] != ':' || content[index + 1] != ':') continue;
            int markerLength = 2;
            while (markerLength < 4 && index + markerLength < content.Length && content[index + markerLength] == ':') markerLength++;
            int after = index + markerLength;
            if (after < content.Length && content[after] != ' ' && content[after] != '\t') continue;
            int descriptionStart = after;
            while (descriptionStart < content.Length && (content[descriptionStart] == ' ' || content[descriptionStart] == '\t')) descriptionStart++;
            parts = new DescriptionListParts(
                content.Substring(0, index),
                content.Substring(index, markerLength),
                content.Substring(descriptionStart),
                index,
                markerLength,
                descriptionStart);
            return true;
        }
        return false;
    }

    internal static bool IsListContinuation(string content) => string.Equals(content, "+", StringComparison.Ordinal);

    internal static bool TryParseListItem(string content, out ListItemParts parts) {
        parts = default;
        if (content.Length < 2) return false;

        char marker = content[0];
        AsciiDocListKind kind;
        int markerLength;
        if (marker == '*') {
            kind = AsciiDocListKind.Unordered;
            markerLength = CountRun(content, '*');
        } else if (marker == '.') {
            kind = AsciiDocListKind.Ordered;
            markerLength = CountRun(content, '.');
        } else if (marker == '-') {
            kind = AsciiDocListKind.Unordered;
            markerLength = 1;
        } else {
            return false;
        }

        if (markerLength >= content.Length || content[markerLength] != ' ') return false;
        parts = new ListItemParts(kind, content.Substring(0, markerLength), markerLength, content.Substring(markerLength + 1));
        return true;
    }

    internal static bool TryParseBlockMacro(string content, out BlockMacroParts parts) {
        parts = default;
        int separator = content.IndexOf("::", StringComparison.Ordinal);
        if (separator <= 0) return false;
        string name = content.Substring(0, separator);
        if (!AsciiDocText.IsMacroName(name)) return false;

        int attributesOpen = content.IndexOf('[', separator + 2);
        if (attributesOpen < separator + 2 || content.Length == 0 || content[content.Length - 1] != ']') return false;

        parts = new BlockMacroParts(
            name,
            content.Substring(separator + 2, attributesOpen - (separator + 2)),
            content.Substring(attributesOpen + 1, content.Length - attributesOpen - 2),
            separator,
            attributesOpen);
        return true;
    }

    internal static bool IsBlockStart(string content) {
        if (IsBlank(content)) return true;
        if (TryGetDelimiter(content, out _)) return true;
        if (IsLineComment(content)) return true;
        if (TryParseHeading(content, out _, out _)) return true;
        if (TryParseAttribute(content, out _)) return true;
        if (TryParseBlockAttributeList(content, out _)) return true;
        if (TryParseBlockTitle(content, out _)) return true;
        if (TryParseBlockAnchor(content, out _, out _)) return true;
        if (TryParseAdmonition(content, out _)) return true;
        if (TryParseDescriptionListItem(content, out _)) return true;
        if (IsListContinuation(content)) return true;
        if (TryParseListItem(content, out _)) return true;
        return TryParseBlockMacro(content, out _);
    }

    private static int CountRun(string content, char marker) {
        int count = 0;
        while (count < content.Length && content[count] == marker) count++;
        return count;
    }

    internal readonly struct AttributeParts {
        internal AttributeParts(string name, string value, bool isUnset, int nameStart, int nameLength, int separator, int valueStart) {
            Name = name;
            Value = value;
            IsUnset = isUnset;
            NameStart = nameStart;
            NameLength = nameLength;
            Separator = separator;
            ValueStart = valueStart;
        }
        internal string Name { get; }
        internal string Value { get; }
        internal bool IsUnset { get; }
        internal int NameStart { get; }
        internal int NameLength { get; }
        internal int Separator { get; }
        internal int ValueStart { get; }
    }

    internal readonly struct ListItemParts {
        internal ListItemParts(AsciiDocListKind kind, string marker, int markerLength, string text) {
            Kind = kind;
            Marker = marker;
            MarkerLength = markerLength;
            Text = text;
        }
        internal AsciiDocListKind Kind { get; }
        internal string Marker { get; }
        internal int MarkerLength { get; }
        internal string Text { get; }
    }

    internal readonly struct BlockMacroParts {
        internal BlockMacroParts(string name, string target, string attributeList, int separator, int attributesOpen) {
            Name = name;
            Target = target;
            AttributeList = attributeList;
            Separator = separator;
            AttributesOpen = attributesOpen;
        }
        internal string Name { get; }
        internal string Target { get; }
        internal string AttributeList { get; }
        internal int Separator { get; }
        internal int AttributesOpen { get; }
    }

    internal readonly struct AdmonitionParts {
        internal AdmonitionParts(AsciiDocAdmonitionKind kind, string label, int textStart, string text) {
            Kind = kind;
            Label = label;
            TextStart = textStart;
            Text = text;
        }
        internal AsciiDocAdmonitionKind Kind { get; }
        internal string Label { get; }
        internal int TextStart { get; }
        internal string Text { get; }
    }

    internal readonly struct DescriptionListParts {
        internal DescriptionListParts(string term, string marker, string description, int markerStart, int markerLength, int descriptionStart) {
            Term = term;
            Marker = marker;
            Description = description;
            MarkerStart = markerStart;
            MarkerLength = markerLength;
            DescriptionStart = descriptionStart;
        }
        internal string Term { get; }
        internal string Marker { get; }
        internal string Description { get; }
        internal int MarkerStart { get; }
        internal int MarkerLength { get; }
        internal int DescriptionStart { get; }
    }
}
