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
                continue;
            }

            var sourceSpan = CreateLineSpan(state, state.SourceLineOffset + idx + 1, state.SourceLineOffset + idx + 1);
            state.Abbreviations[label] = new MarkdownAbbreviationDefinition(
                label,
                title,
                sourceSpan,
                labelSpan,
                titleSpan,
                openingMarkerSpan,
                separatorMarkerSpan);
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
        out MarkdownSourceSpan? separatorMarkerSpan) {
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

        if (titleStart >= trimmed.Length) {
            return false;
        }

        label = trimmed.Substring(2, closingBracket - 2);
        title = trimmed.Substring(titleStart);
        if (string.IsNullOrWhiteSpace(label) || string.IsNullOrWhiteSpace(title)) {
            return false;
        }

        int absoluteLine = state?.SourceLineOffset + lineIndex + 1 ?? lineIndex + 1;
        labelSpan = CreateSpan(state, absoluteLine, leading + 3, absoluteLine, leading + closingBracket);
        titleSpan = CreateSpan(state, absoluteLine, leading + titleStart + 1, absoluteLine, leading + titleStart + title.Length);
        openingMarkerSpan = CreateSpan(state, absoluteLine, leading + 1, absoluteLine, leading + 2);
        separatorMarkerSpan = CreateSpan(state, absoluteLine, leading + closingBracket + 1, absoluteLine, leading + closingBracket + 2);
        return true;
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

            if (!IsAbbreviationBoundary(text, position - 1) || !IsAbbreviationBoundary(text, position + candidate.Label.Length)) {
                continue;
            }

            definition = candidate;
            return true;
        }

        return false;
    }

    private static bool IsAbbreviationBoundary(string text, int index) {
        if (index < 0 || index >= text.Length) {
            return true;
        }

        var value = text[index];
        return !char.IsLetterOrDigit(value) && value != '-';
    }
}
