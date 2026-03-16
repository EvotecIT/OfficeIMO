using System.Globalization;

namespace OfficeIMO.Markdown;

internal static class MarkdownInlineTransformHelpers {
    internal static InlineSequence CreateSequence(IEnumerable<IMarkdownInline> nodes) {
        var sequence = new InlineSequence {
            AutoSpacing = false
        };

        if (nodes == null) {
            return sequence;
        }

        foreach (var node in nodes) {
            if (node != null) {
                sequence.AddRaw(node);
            }
        }

        return sequence;
    }

    internal static InlineSequence TrimWhitespace(InlineSequence source, bool trimStart, bool trimEnd) {
        if (source == null || source.Nodes.Count == 0 || (!trimStart && !trimEnd)) {
            return source ?? CreateSequence(Array.Empty<IMarkdownInline>());
        }

        var nodes = new List<IMarkdownInline>(source.Nodes);
        bool changed = false;

        if (trimStart) {
            while (nodes.Count > 0 && nodes[0] is TextRun leading) {
                var trimmed = leading.Text.TrimStart();
                if (trimmed == leading.Text) {
                    break;
                }

                changed = true;
                if (trimmed.Length == 0) {
                    nodes.RemoveAt(0);
                    continue;
                }

                nodes[0] = new TextRun(trimmed);
                break;
            }
        }

        if (trimEnd) {
            while (nodes.Count > 0 && nodes[nodes.Count - 1] is TextRun trailing) {
                var trimmed = trailing.Text.TrimEnd();
                if (trimmed == trailing.Text) {
                    break;
                }

                changed = true;
                if (trimmed.Length == 0) {
                    nodes.RemoveAt(nodes.Count - 1);
                    continue;
                }

                nodes[nodes.Count - 1] = new TextRun(trimmed);
                break;
            }
        }

        return changed ? CreateSequence(nodes) : source;
    }

    internal static bool HasVisibleContent(InlineSequence? sequence) {
        var nodes = sequence?.Nodes;
        if (nodes == null || nodes.Count == 0) {
            return false;
        }

        for (var i = 0; i < nodes.Count; i++) {
            var node = nodes[i];
            switch (node) {
                case null:
                    continue;
                case TextRun textRun when string.IsNullOrWhiteSpace(textRun.Text):
                    continue;
                default:
                    return true;
            }
        }

        return false;
    }

    internal static bool StartsWithStrong(InlineSequence? sequence) {
        var nodes = sequence?.Nodes;
        if (nodes == null) {
            return false;
        }

        for (var i = 0; i < nodes.Count; i++) {
            var node = nodes[i];
            switch (node) {
                case null:
                    continue;
                case TextRun textRun when string.IsNullOrWhiteSpace(textRun.Text):
                    continue;
                default:
                    return node is IStrongMarkdownInline;
            }
        }

        return false;
    }

    internal static IReadOnlyList<ListItem> ExpandCompactStrongLabelListItems(
        InlineSequence source,
        int level,
        bool forceLoose) {
        var items = new List<ListItem>();
        var current = TrimWhitespace(source, trimStart: true, trimEnd: true);

        while (TrySplitCompactStrongLabelBoundary(current, out var head, out _, out var tail)) {
            head = TrimWhitespace(head, trimStart: true, trimEnd: true);
            tail = TrimWhitespace(tail, trimStart: true, trimEnd: true);
            if (!HasVisibleContent(head) || !StartsWithStrong(tail)) {
                break;
            }

            items.Add(CreatePlainListItem(head, level, forceLoose));
            current = tail;
        }

        current = TrimWhitespace(current, trimStart: true, trimEnd: true);
        if (HasVisibleContent(current)) {
            items.Add(CreatePlainListItem(current, level, forceLoose));
        }

        return items;
    }

    internal static ListItem CreatePlainListItem(InlineSequence content, int level, bool forceLoose) {
        return new ListItem(content) {
            Level = level,
            ForceLoose = forceLoose
        };
    }

    internal static bool TrySplitCompactHeadingBoundary(
        InlineSequence source,
        out InlineSequence head,
        out int level,
        out InlineSequence tail) {
        level = 0;
        return TrySplitAtBoundary(source, TryMatchCompactHeadingBoundary, out head, out var marker, out tail)
               && int.TryParse(marker, NumberStyles.Integer, CultureInfo.InvariantCulture, out level);
    }

    internal static bool TrySplitHeadingListBoundary(
        InlineSequence source,
        out InlineSequence head,
        out char marker,
        out InlineSequence tail) {
        if (!TrySplitAtBoundary(source, TryMatchHeadingListBoundary, out head, out var token, out tail)
            || token.Length != 1) {
            marker = default;
            return false;
        }

        marker = token[0];
        return true;
    }

    internal static bool TrySplitCompactStrongLabelBoundary(
        InlineSequence source,
        out InlineSequence head,
        out char marker,
        out InlineSequence tail) {
        if (!TrySplitAtBoundary(source, TryMatchCompactStrongLabelBoundary, out head, out var token, out tail)
            || token.Length != 1) {
            marker = default;
            return false;
        }

        marker = token[0];
        return true;
    }

    internal static bool TrySplitColonListBoundary(
        InlineSequence source,
        out InlineSequence head,
        out char marker,
        out InlineSequence tail) {
        if (!TrySplitAtBoundary(source, TryMatchColonListBoundary, out head, out var token, out tail)
            || token.Length != 1) {
            marker = default;
            return false;
        }

        marker = token[0];
        return true;
    }

    private static bool TrySplitAtBoundary(
        InlineSequence source,
        Func<string, bool, char, BoundaryMatch?> matcher,
        out InlineSequence head,
        out string marker,
        out InlineSequence tail) {
        head = CreateSequence(Array.Empty<IMarkdownInline>());
        tail = CreateSequence(Array.Empty<IMarkdownInline>());
        marker = string.Empty;

        var nodes = source?.Nodes;
        if (nodes == null || nodes.Count == 0) {
            return false;
        }

        var headNodes = new List<IMarkdownInline>(nodes.Count);
        bool hasPreviousChar = false;
        char previousChar = '\0';

        for (var i = 0; i < nodes.Count; i++) {
            var node = nodes[i];
            if (node is TextRun textRun) {
                var match = matcher(textRun.Text, hasPreviousChar, previousChar);
                if (match != null) {
                    AppendText(headNodes, textRun.Text.Substring(0, match.StartIndex));

                    var tailNodes = new List<IMarkdownInline>(nodes.Count - i);
                    AppendText(tailNodes, textRun.Text.Substring(match.EndIndex));
                    for (var j = i + 1; j < nodes.Count; j++) {
                        if (nodes[j] != null) {
                            tailNodes.Add(nodes[j]);
                        }
                    }

                    head = CreateSequence(headNodes);
                    tail = CreateSequence(tailNodes);
                    marker = match.Marker;
                    return true;
                }
            }

            if (node != null) {
                headNodes.Add(node);
            }

            if (TryGetMaskedTrailingChar(node, out var trailingChar)) {
                hasPreviousChar = true;
                previousChar = trailingChar;
            }
        }

        return false;
    }

    private static BoundaryMatch? TryMatchCompactHeadingBoundary(string text, bool hasPreviousChar, char previousChar) {
        if (string.IsNullOrEmpty(text) || text.IndexOf('#') < 0) {
            return null;
        }

        for (var i = 0; i < text.Length; i++) {
            if (text[i] != '#') {
                continue;
            }

            var level = 1;
            while (i + level < text.Length && text[i + level] == '#' && level < 6) {
                level++;
            }

            if (level < 2) {
                continue;
            }

            char before = i > 0 ? text[i - 1] : hasPreviousChar ? previousChar : '\0';
            if (before == '\0' || char.IsWhiteSpace(before)) {
                continue;
            }

            var j = i + level;
            if (j >= text.Length || !char.IsWhiteSpace(text[j])) {
                continue;
            }

            while (j < text.Length && char.IsWhiteSpace(text[j])) {
                j++;
            }

            return new BoundaryMatch(i, j, level.ToString(CultureInfo.InvariantCulture));
        }

        return null;
    }

    private static BoundaryMatch? TryMatchHeadingListBoundary(string text, bool hasPreviousChar, char previousChar) {
        if (string.IsNullOrEmpty(text)) {
            return null;
        }

        for (var i = 0; i < text.Length; i++) {
            var marker = text[i];
            if (marker != '-' && marker != '+' && marker != '*') {
                continue;
            }

            char before = i > 0 ? text[i - 1] : hasPreviousChar ? previousChar : '\0';
            if (before == '\0' || char.IsWhiteSpace(before)) {
                continue;
            }

            var j = i + 1;
            if (j >= text.Length || !char.IsWhiteSpace(text[j])) {
                continue;
            }

            while (j < text.Length && char.IsWhiteSpace(text[j])) {
                j++;
            }

            return new BoundaryMatch(i, j, marker.ToString());
        }

        return null;
    }

    private static BoundaryMatch? TryMatchCompactStrongLabelBoundary(string text, bool hasPreviousChar, char previousChar) {
        if (string.IsNullOrEmpty(text)) {
            return null;
        }

        for (var i = 0; i < text.Length; i++) {
            var marker = text[i];
            if (marker != '-' && marker != '+' && marker != '*') {
                continue;
            }

            char before = i > 0 ? text[i - 1] : hasPreviousChar ? previousChar : '\0';
            if (!IsCompactListBoundaryPrefix(before)) {
                continue;
            }

            var j = i + 1;
            if (j >= text.Length || !char.IsWhiteSpace(text[j])) {
                continue;
            }

            while (j < text.Length && char.IsWhiteSpace(text[j])) {
                j++;
            }

            return new BoundaryMatch(i, j, marker.ToString());
        }

        return null;
    }

    private static BoundaryMatch? TryMatchColonListBoundary(string text, bool hasPreviousChar, char previousChar) {
        if (string.IsNullOrEmpty(text) || text.IndexOf(':') < 0) {
            return null;
        }

        for (var i = 0; i < text.Length; i++) {
            if (text[i] != ':') {
                continue;
            }

            var j = i + 1;
            while (j < text.Length && char.IsWhiteSpace(text[j])) {
                j++;
            }

            if (j >= text.Length) {
                continue;
            }

            var marker = text[j];
            if (marker != '-' && marker != '+' && marker != '*') {
                continue;
            }

            j++;
            if (j >= text.Length || !char.IsWhiteSpace(text[j])) {
                continue;
            }

            while (j < text.Length && char.IsWhiteSpace(text[j])) {
                j++;
            }

            return new BoundaryMatch(i + 1, j, marker.ToString());
        }

        return null;
    }

    private static bool TryGetMaskedTrailingChar(IMarkdownInline? node, out char ch) {
        ch = '\0';
        if (node == null) {
            return false;
        }

        if (node is CodeSpanInline) {
            ch = ' ';
            return true;
        }

        if (node is TextRun textRun) {
            if (textRun.Text.Length == 0) {
                return false;
            }

            ch = textRun.Text[textRun.Text.Length - 1];
            return true;
        }

        var plain = new System.Text.StringBuilder();
        ((IPlainTextMarkdownInline)node).AppendPlainText(plain);
        if (plain.Length == 0) {
            return false;
        }

        ch = plain[plain.Length - 1];
        return true;
    }

    private static bool IsCompactListBoundaryPrefix(char ch) {
        if (ch == '\0') {
            return false;
        }

        return ch == ')' || char.IsPunctuation(ch) || char.IsSymbol(ch);
    }

    private static void AppendText(List<IMarkdownInline> nodes, string? text) {
        if (!string.IsNullOrEmpty(text)) {
            nodes.Add(new TextRun(text!));
        }
    }

    private sealed class BoundaryMatch {
        public BoundaryMatch(int startIndex, int endIndex, string marker) {
            StartIndex = startIndex;
            EndIndex = endIndex;
            Marker = marker;
        }

        public int StartIndex { get; }
        public int EndIndex { get; }
        public string Marker { get; }
    }
}
