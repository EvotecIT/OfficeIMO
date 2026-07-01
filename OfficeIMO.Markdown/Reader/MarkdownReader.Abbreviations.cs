namespace OfficeIMO.Markdown;

public static partial class MarkdownReader {
    internal sealed class AbbreviationDefParser : IMarkdownBlockParser {
        public bool TryParse(string[] lines, ref int i, MarkdownReaderOptions options, MarkdownDoc doc, MarkdownReaderState state) {
            if (options?.Abbreviations != true || i < 0 || i >= lines.Length) {
                return false;
            }

            if (!TryParseAbbreviationDefinition(
                lines[i],
                i,
                state,
                out var label,
                out var title,
                out var labelSpan,
                out var titleSpan,
                out var openingMarkerSpan,
                out var separatorMarkerSpan)) {
                return false;
            }

            var sourceSpan = CreateLineSpan(state, state.SourceLineOffset + i + 1, state.SourceLineOffset + i + 1);
            state.Abbreviations[label] = new MarkdownAbbreviationDefinition(
                label,
                title,
                sourceSpan,
                labelSpan,
                titleSpan,
                openingMarkerSpan,
                separatorMarkerSpan);
            i++;
            return true;
        }
    }

    private static IReadOnlyList<MarkdownAbbreviationDefinition> SnapshotAbbreviationDefinitions(MarkdownReaderState state) {
        if (state == null || state.Abbreviations.Count == 0) {
            return Array.Empty<MarkdownAbbreviationDefinition>();
        }

        return state.Abbreviations.Values
            .OrderBy(definition => definition.LabelSourceSpan?.StartLine ?? int.MaxValue)
            .ThenBy(definition => definition.LabelSourceSpan?.StartColumn ?? int.MaxValue)
            .ThenBy(definition => definition.Label, StringComparer.Ordinal)
            .ToArray();
    }

    private static void PreScanAbbreviationDefinitions(string[] lines, MarkdownReaderState state, MarkdownReaderOptions options) {
        if (options?.Abbreviations != true || lines == null || lines.Length == 0) {
            return;
        }

        bool inFence = false;
        char fenceChar = '\0';
        int fenceLen = 0;

        for (int idx = 0; idx < lines.Length; idx++) {
            var line = lines[idx];
            if (string.IsNullOrWhiteSpace(line)) {
                continue;
            }

            if (!inFence) {
                if (IsCodeFenceOpen(line, out _, out fenceChar, out fenceLen)) {
                    inFence = true;
                    continue;
                }
            } else {
                if (IsCodeFenceClose(line, fenceChar, fenceLen)) {
                    inFence = false;
                }
                continue;
            }

            var isListItemDefinition = false;
            if (!TryParseAbbreviationDefinition(
                    line,
                    idx,
                    state,
                    out var label,
                    out var title,
                    out var labelSpan,
                    out var titleSpan,
                    out var openingMarkerSpan,
                    out var separatorMarkerSpan)) {
                if (!TryParseListItemAbbreviationDefinition(
                    line,
                    idx,
                    state,
                    options,
                    out label,
                    out title,
                    out labelSpan,
                    out titleSpan,
                    out openingMarkerSpan,
                    out separatorMarkerSpan)) {
                    continue;
                }

                isListItemDefinition = true;
            }

            var sourceSpan = CreateLineSpan(state, state.SourceLineOffset + idx + 1, state.SourceLineOffset + idx + 1);
            state.Abbreviations[label] = new MarkdownAbbreviationDefinition(
                label,
                title,
                sourceSpan,
                labelSpan,
                titleSpan,
                openingMarkerSpan,
                separatorMarkerSpan,
                isListItemDefinition);
        }
    }

    private static bool TryParseAbbreviationDefinition(
        string line,
        int lineIndex,
        MarkdownReaderState? state,
        out string label,
        out string title,
        out MarkdownSourceSpan? labelSpan,
        out MarkdownSourceSpan? titleSpan,
        out MarkdownSourceSpan? openingMarkerSpan,
        out MarkdownSourceSpan? separatorMarkerSpan,
        int columnOffset = 0) {
        label = string.Empty;
        title = string.Empty;
        labelSpan = null;
        titleSpan = null;
        openingMarkerSpan = null;
        separatorMarkerSpan = null;

        if (string.IsNullOrWhiteSpace(line)) {
            return false;
        }

        int leading = 0;
        while (leading < line.Length && line[leading] == ' ') {
            leading++;
        }

        if (leading >= 4 || (leading < line.Length && line[leading] == '\t')) {
            return false;
        }

        var trimmed = line.Substring(leading).TrimEnd();
        if (trimmed.Length < 6 || trimmed[0] != '*' || trimmed[1] != '[') {
            return false;
        }

        int closingBracket = trimmed.IndexOf(']', 2);
        if (closingBracket <= 2 || closingBracket + 1 >= trimmed.Length || trimmed[closingBracket + 1] != ':') {
            return false;
        }

        int titleStart = closingBracket + 2;
        while (titleStart < trimmed.Length && char.IsWhiteSpace(trimmed[titleStart])) {
            titleStart++;
        }

        label = trimmed.Substring(2, closingBracket - 2);
        title = titleStart < trimmed.Length ? trimmed.Substring(titleStart) : string.Empty;
        if (string.IsNullOrWhiteSpace(label)) {
            return false;
        }

        int absoluteLine = state?.SourceLineOffset + lineIndex + 1 ?? lineIndex + 1;
        labelSpan = CreateSpan(state, absoluteLine, columnOffset + leading + 3, absoluteLine, columnOffset + leading + closingBracket);
        titleSpan = title.Length > 0
            ? CreateSpan(state, absoluteLine, columnOffset + leading + titleStart + 1, absoluteLine, columnOffset + leading + titleStart + title.Length)
            : null;
        openingMarkerSpan = CreateSpan(state, absoluteLine, columnOffset + leading + 1, absoluteLine, columnOffset + leading + 2);
        separatorMarkerSpan = CreateSpan(state, absoluteLine, columnOffset + leading + closingBracket + 1, absoluteLine, columnOffset + leading + closingBracket + 2);
        return true;
    }

    private static bool TryParseListItemAbbreviationDefinition(
        string line,
        int lineIndex,
        MarkdownReaderState? state,
        MarkdownReaderOptions? options,
        out string label,
        out string title,
        out MarkdownSourceSpan? labelSpan,
        out MarkdownSourceSpan? titleSpan,
        out MarkdownSourceSpan? openingMarkerSpan,
        out MarkdownSourceSpan? separatorMarkerSpan) {

        label = string.Empty;
        title = string.Empty;
        labelSpan = null;
        titleSpan = null;
        openingMarkerSpan = null;
        separatorMarkerSpan = null;

        if (string.IsNullOrEmpty(line) || !TryGetRawListItemContentAfterMarker(line, out var content, options)) {
            return false;
        }

        int contentStartIndex = line.Length - content.Length;
        return TryParseAbbreviationDefinition(
            content,
            lineIndex,
            state,
            out label,
            out title,
            out labelSpan,
            out titleSpan,
            out openingMarkerSpan,
            out separatorMarkerSpan,
            contentStartIndex);
    }

    private static bool IsAbbreviationDefinitionLine(string line) {
        return TryParseAbbreviationDefinition(
            line,
            0,
            null,
            out _,
            out _,
            out _,
            out _,
            out _,
            out _);
    }

    private static bool TryConsumeAbbreviation(
        string text,
        int position,
        MarkdownReaderState? state,
        out MarkdownAbbreviationDefinition definition) {
        definition = null!;
        if (string.IsNullOrEmpty(text) || state == null || state.Abbreviations.Count == 0 || position < 0 || position >= text.Length) {
            return false;
        }

        foreach (var candidate in state.Abbreviations.Values.OrderByDescending(static item => item.Label.Length)) {
            if (candidate.Label.Length == 0 || position + candidate.Label.Length > text.Length) {
                continue;
            }

            if (!string.Equals(text.Substring(position, candidate.Label.Length), candidate.Label, StringComparison.Ordinal)) {
                continue;
            }

            if (!IsAbbreviationOpeningBoundary(text, position - 1)
                || !IsAbbreviationClosingBoundary(text, position + candidate.Label.Length)) {
                continue;
            }

            definition = candidate;
            return true;
        }

        return false;
    }

    private static bool ContainsAbbreviationCandidate(string text, MarkdownReaderState? state) {
        if (string.IsNullOrEmpty(text) || state == null || state.Abbreviations.Count == 0) {
            return false;
        }

        for (int position = 0; position < text.Length; position++) {
            if (TryConsumeAbbreviation(text, position, state, out _)) {
                return true;
            }
        }

        return false;
    }

    private static bool IsAbbreviationOpeningBoundary(string text, int index) {
        if (index < 0 || index >= text.Length) {
            return true;
        }

        var value = text[index];
        return char.IsWhiteSpace(value) || value == '_' || value == '[' || value == '*';
    }

    private static bool IsAbbreviationClosingBoundary(string text, int index) {
        if (index < 0 || index >= text.Length) {
            return true;
        }

        var value = text[index];
        if (char.IsLetterOrDigit(value)) {
            return false;
        }

        if (value != '-') {
            return true;
        }

        int next = index + 1;
        return next >= text.Length || !char.IsLetterOrDigit(text[next]);
    }
}
