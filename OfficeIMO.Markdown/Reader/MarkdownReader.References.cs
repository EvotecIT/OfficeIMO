using System.IO;
using System.Linq;
using System.Text;
// Intentionally avoid heavy regex use; simple scanning is used for resilience and speed.

namespace OfficeIMO.Markdown;

public static partial class MarkdownReader {
    private static void PreScanReferenceLinkDefinitions(string[] lines, MarkdownReaderState state) {
        PreScanReferenceLinkDefinitions(lines, state, new MarkdownReaderOptions());
    }

    private static IReadOnlyList<MarkdownReferenceLinkDefinition> SnapshotReferenceLinkDefinitions(MarkdownReaderState state) {
        if (state == null || state.LinkRefs.Count == 0) {
            return Array.Empty<MarkdownReferenceLinkDefinition>();
        }

        return state.LinkRefs.Values
            .OrderBy(definition => definition.LabelSourceSpan?.StartLine ?? int.MaxValue)
            .ThenBy(definition => definition.LabelSourceSpan?.StartColumn ?? int.MaxValue)
            .ThenBy(definition => definition.Label, StringComparer.Ordinal)
            .ToArray();
    }

    private static void PreScanReferenceLinkDefinitions(string[] lines, MarkdownReaderState state, MarkdownReaderOptions options) {
        bool inFence = false;
        char fenceChar = '\0';
        int fenceLen = 0;
        bool openParagraph = false;
        bool inQuotedFence = false;
        char quotedFenceChar = '\0';
        int quotedFenceLen = 0;

        for (int idx = 0; idx < lines.Length; idx++) {
            var line = lines[idx];
            if (string.IsNullOrWhiteSpace(line)) {
                openParagraph = false;
                continue;
            }

            if (TryUpdateQuotedFencePreScanState(line, ref inQuotedFence, ref quotedFenceChar, ref quotedFenceLen)) {
                openParagraph = false;
                continue;
            }

            // Ignore anything inside fenced code blocks.
            if (!inFence) {
                if (IsCodeFenceOpen(line, out _, out fenceChar, out fenceLen)) {
                    inFence = true;
                    openParagraph = false;
                    continue;
                }
            } else {
                if (IsCodeFenceClose(line, fenceChar, fenceLen)) {
                    inFence = false;
                }
                continue;
            }

            // Ignore indented code blocks (4+ leading spaces or a tab). Reference definitions are only valid
            // up to 3 leading spaces in typical Markdown implementations.
            int leading = 0;
            while (leading < line.Length && line[leading] == ' ') leading++;
            if (leading >= 4) continue;
            if (leading < line.Length && line[leading] == '\t') continue;

            if (TryParseReferenceLinkDefinition(
                lines,
                idx,
                options,
                state,
                out var label,
                out var url,
                out var title,
                out var consumedLines,
                out var labelSpan,
                out var urlSpan,
                out var titleSpan,
                out var openingMarkerSpan,
                out var separatorMarkerSpan)) {
                if (openParagraph) {
                    continue;
                }

                var resolved = ResolveUrl(url, options);
                if (resolved != null && !state.LinkRefs.ContainsKey(label)) {
                    var sourceSpan = CreateLineSpan(
                        state,
                        state.SourceLineOffset + idx + 1,
                        state.SourceLineOffset + idx + consumedLines);
                    state.LinkRefs[label] = new MarkdownReferenceLinkDefinition(
                        label,
                        resolved!,
                        title,
                        sourceSpan,
                        labelSpan,
                        urlSpan,
                        titleSpan,
                        openingMarkerSpan,
                        separatorMarkerSpan);
                }
                idx += consumedLines - 1;
                openParagraph = false;
                continue;
            }

            if (!inQuotedFence && TryParseQuotedReferenceLinkDefinition(
                lines,
                idx,
                options,
                state,
                out var quotedDefinition,
                out var quotedConsumedLines)) {
                if (!state.LinkRefs.ContainsKey(quotedDefinition.Label)) {
                    state.LinkRefs[quotedDefinition.Label] = quotedDefinition;
                }

                idx += quotedConsumedLines - 1;
                openParagraph = false;
                continue;
            }

            openParagraph = IsReferenceDefinitionParagraphContinuationLine(lines, idx, options);
        }
    }

    private static void CaptureConsumedSyntaxNodes(
        IMarkdownBlockParser parser,
        string[] lines,
        int startIndex,
        MarkdownReaderOptions options,
        List<MarkdownSyntaxNode> syntaxNodes,
        MarkdownReaderState state) {
        if (parser is not ReferenceLinkDefParser) {
            return;
        }

        if (TryBuildReferenceDefinitionSyntaxNode(lines, startIndex, options, state, out var node, out var consumedLines)) {
            syntaxNodes.Add(node);
        }
    }

    private static bool TryBuildReferenceDefinitionSyntaxNode(
        string[] lines,
        int index,
        MarkdownReaderOptions options,
        MarkdownReaderState state,
        out MarkdownSyntaxNode node,
        out int consumedLines) {
        node = null!;
        consumedLines = 0;

        if (!TryParseReferenceLinkDefinition(
            lines,
            index,
            options,
            state,
            out var label,
            out var url,
            out var title,
            out consumedLines,
            out var labelSpan,
            out var urlSpan,
            out var titleSpan,
            out var openingMarkerSpan,
            out var separatorMarkerSpan)) {
            return false;
        }

        var children = new List<MarkdownSyntaxNode>(5);
        AddReferenceDefinitionLabelFrameChildren(children, labelSpan, label, openingMarkerSpan, separatorMarkerSpan);

        if (!string.IsNullOrEmpty(url)) {
            children.Add(new MarkdownSyntaxNode(MarkdownSyntaxKind.ReferenceLinkUrl, urlSpan, url));
        }

        if (!string.IsNullOrEmpty(title)) {
            children.Add(new MarkdownSyntaxNode(MarkdownSyntaxKind.ReferenceLinkTitle, titleSpan, title));
        }

        var definitionSpan = CreateLineSpan(
            state,
            state.SourceLineOffset + index + 1,
            state.SourceLineOffset + index + consumedLines);
        var literal = consumedLines > 1
            ? string.Join("\n", lines.Skip(index).Take(consumedLines))
            : lines[index];

        node = new MarkdownSyntaxNode(
            MarkdownSyntaxKind.ReferenceLinkDefinition,
            definitionSpan,
            literal,
            children);
        return true;
    }

    private static void AddReferenceDefinitionLabelFrameChildren(
        List<MarkdownSyntaxNode> children,
        MarkdownSourceSpan? labelSpan,
        string label,
        MarkdownSourceSpan? openingMarkerSpan,
        MarkdownSourceSpan? separatorMarkerSpan) {
        if (openingMarkerSpan.HasValue && separatorMarkerSpan.HasValue) {
            children.Add(new MarkdownSyntaxNode(MarkdownSyntaxKind.ReferenceLinkOpeningMarker, openingMarkerSpan, "["));
            children.Add(new MarkdownSyntaxNode(MarkdownSyntaxKind.ReferenceLinkLabel, labelSpan, label));
            children.Add(new MarkdownSyntaxNode(MarkdownSyntaxKind.ReferenceLinkSeparatorMarker, separatorMarkerSpan, "]:"));
            return;
        }

        children.Add(new MarkdownSyntaxNode(MarkdownSyntaxKind.ReferenceLinkLabel, labelSpan, label));
    }

    private static bool TryParseReferenceLinkDefinition(
        string[] lines,
        int index,
        MarkdownReaderOptions options,
        MarkdownReaderState? state,
        out string label,
        out string url,
        out string? title,
        out int consumedLines,
        out MarkdownSourceSpan? labelSpan,
        out MarkdownSourceSpan? urlSpan,
        out MarkdownSourceSpan? titleSpan,
        out MarkdownSourceSpan? openingMarkerSpan,
        out MarkdownSourceSpan? separatorMarkerSpan) {
        label = url = string.Empty;
        title = null;
        consumedLines = 0;
        labelSpan = null;
        urlSpan = null;
        titleSpan = null;
        openingMarkerSpan = null;
        separatorMarkerSpan = null;

        if (index < 0 || index >= lines.Length) return false;
        if (!TryParseReferenceDefinitionLabel(
            lines,
            index,
            state,
            out label,
            out var rest,
            out var labelConsumedLines,
            out labelSpan,
            out openingMarkerSpan,
            out separatorMarkerSpan,
            out var restLineIndex,
            out var restStartColumnZeroBased)) {
            return false;
        }

        if (string.IsNullOrEmpty(rest)) {
            if (!TryParseReferenceDestinationContinuation(
                lines,
                restLineIndex + 1,
                out rest,
                out var continuationOffset,
                out var continuationLeading)) {
                return false;
            }

            int destinationIndex = restLineIndex + continuationOffset;
            if (!TryParseReferenceDestinationAndTitle(
                rest,
                state,
                destinationIndex,
                continuationLeading,
                out url,
                out title,
                out urlSpan,
                out titleSpan)) {
                if (!TryParseReferenceDestinationAndMultilineTitle(
                    lines,
                    destinationIndex,
                    rest,
                    continuationLeading,
                    state,
                    out url,
                    out title,
                    out urlSpan,
                    out titleSpan,
                    out var multilineTitleEndIndex)) {
                    return false;
                }

                consumedLines = multilineTitleEndIndex - index + 1;
                return !string.IsNullOrEmpty(label);
            }

            consumedLines = destinationIndex - index + 1;
            if (title == null && TryParseReferenceTitleContinuation(lines, destinationIndex + 1, state, out var continuedTitle, out var continuedTitleSpan)) {
                title = continuedTitle;
                titleSpan = continuedTitleSpan;
                consumedLines++;
            }

            return !string.IsNullOrEmpty(label);
        }

        if (!TryParseReferenceDestinationAndTitle(
            rest,
            state,
            restLineIndex,
            restStartColumnZeroBased,
            out url,
            out title,
            out urlSpan,
            out titleSpan)) {
            if (!TryParseReferenceDestinationAndMultilineTitle(
                lines,
                restLineIndex,
                rest,
                restStartColumnZeroBased,
                state,
                out url,
                out title,
                out urlSpan,
                out titleSpan,
                out var multilineTitleEndIndex)) {
                return false;
            }

            consumedLines = multilineTitleEndIndex - index + 1;
            return !string.IsNullOrEmpty(label);
        }

        consumedLines = restLineIndex - index + 1;
        if (title == null && TryParseReferenceTitleContinuation(lines, restLineIndex + 1, state, out var continuationTitle, out var continuationTitleSpan)) {
            title = continuationTitle;
            titleSpan = continuationTitleSpan;
            consumedLines++;
        }

        return !string.IsNullOrEmpty(label);
    }

    private static bool TryParseReferenceLinkDefinition(string[] lines, int index, MarkdownReaderOptions options, out string label, out string url, out string? title, out int consumedLines) =>
        TryParseReferenceLinkDefinition(lines, index, options, state: null, out label, out url, out title, out consumedLines, out _, out _, out _, out _, out _);

    private static bool TryParseReferenceTitleContinuation(string[] lines, int index, out string? title) =>
        TryParseReferenceTitleContinuation(lines, index, state: null, out title, out _);

    private static bool TryParseReferenceTitleContinuation(
        string[] lines,
        int index,
        MarkdownReaderState? state,
        out string? title,
        out MarkdownSourceSpan? titleSpan) {
        title = null;
        titleSpan = null;
        if (index < 0 || index >= lines.Length) return false;

        var line = lines[index];
        if (string.IsNullOrWhiteSpace(line)) return false;

        int leading = 0;
        while (leading < line.Length && char.IsWhiteSpace(line[leading])) leading++;
        if (leading >= line.Length) return false;

        var trimmed = line.Substring(leading).TrimEnd();
        if (!TryParseOptionalTitleToken(trimmed, 0, trimmed.Length, out title, out int titleStart, out int titleLength) || title == null) {
            return false;
        }

        title = DecodeLinkDestinationOrTitle(title);
        titleSpan = CreateSpan(
            state,
            state?.SourceLineOffset + index + 1 ?? index + 1,
            leading + titleStart + 1,
            state?.SourceLineOffset + index + 1 ?? index + 1,
            leading + titleStart + titleLength);
        return true;
    }

    private static bool StartsWithReferenceDefinitionLikeLabel(string line) {
        if (string.IsNullOrWhiteSpace(line)) return false;
        var trimmed = line.TrimStart();
        if (trimmed.Length < 4 || trimmed[0] != '[') return false;
        if (trimmed.Length > 1 && trimmed[1] == '^') return false;

        int balancedEnd = FindMatchingBracket(trimmed, 0);
        return balancedEnd >= 1 && balancedEnd + 1 < trimmed.Length && trimmed[balancedEnd + 1] == ':';
    }

    private static string NormalizeReferenceLabel(string? label) {
        if (string.IsNullOrWhiteSpace(label)) return string.Empty;
        var t = label!.Trim();
        var sb = new System.Text.StringBuilder(t.Length);
        bool prevSpace = false;
        for (int i = 0; i < t.Length; i++) {
            char c = t[i];
            if (char.IsWhiteSpace(c)) {
                if (!prevSpace) sb.Append(' ');
                prevSpace = true;
                continue;
            }

            if (char.IsHighSurrogate(c) && i + 1 < t.Length && char.IsLowSurrogate(t[i + 1])) {
                AppendUnicodeCaseFold(sb, t.Substring(i, 2));
                i++;
                prevSpace = false;
                continue;
            }

            AppendUnicodeCaseFold(sb, c.ToString());
            prevSpace = false;
        }
        return sb.ToString();
    }

    private static void AppendUnicodeCaseFold(System.Text.StringBuilder builder, string scalar) {
        switch (scalar) {
            case "ß":
            case "ẞ":
                builder.Append("ss");
                return;
            case "ς":
                builder.Append("σ");
                return;
            default:
                builder.Append(scalar.ToLowerInvariant());
                return;
        }
    }

    private static bool TryParseReferenceDefinitionLabel(
        string[] lines,
        int index,
        MarkdownReaderState? state,
        out string label,
        out string rest,
        out int consumedLines,
        out MarkdownSourceSpan? labelSpan,
        out MarkdownSourceSpan? openingMarkerSpan,
        out MarkdownSourceSpan? separatorMarkerSpan,
        out int restLineIndex,
        out int restStartColumnZeroBased) {
        label = string.Empty;
        rest = string.Empty;
        consumedLines = 0;
        labelSpan = null;
        openingMarkerSpan = null;
        separatorMarkerSpan = null;
        restLineIndex = index;
        restStartColumnZeroBased = 0;

        if (index < 0 || index >= lines.Length) {
            return false;
        }

        var line = lines[index];
        if (string.IsNullOrWhiteSpace(line)) {
            return false;
        }

        int leading = 0;
        while (leading < line.Length && line[leading] == ' ') {
            leading++;
        }

        if (leading >= 4) {
            return false;
        }

        if (leading < line.Length && line[leading] == '\t') {
            return false;
        }

        var trimmed = line.Substring(leading).TrimEnd();
        if (trimmed.Length < 1 || trimmed[0] != '[') {
            return false;
        }

        if (trimmed.Length > 1 && trimmed[1] == '^') {
            return false;
        }

        int rb = FindReferenceLabelEnd(trimmed, 0);
        if (rb > 1 && rb + 1 < trimmed.Length && trimmed[rb + 1] == ':') {
            label = NormalizeReferenceLabel(trimmed.Substring(1, rb - 1));
            int absoluteLine = state?.SourceLineOffset + index + 1 ?? index + 1;
            labelSpan = CreateSpan(
                state,
                absoluteLine,
                leading + 2,
                absoluteLine,
                leading + rb);
            openingMarkerSpan = CreateSpan(state, absoluteLine, leading + 1, absoluteLine, leading + 1);
            separatorMarkerSpan = CreateSpan(state, absoluteLine, leading + rb + 1, absoluteLine, leading + rb + 2);

            int restStart = rb + 2;
            while (restStart < trimmed.Length && char.IsWhiteSpace(trimmed[restStart])) {
                restStart++;
            }

            rest = trimmed.Substring(restStart);
            consumedLines = 1;
            restLineIndex = index;
            restStartColumnZeroBased = leading + restStart;
            return !string.IsNullOrEmpty(label);
        }

        if (rb >= 0) {
            return false;
        }

        var labelBuilder = new System.Text.StringBuilder(trimmed.Substring(1));
        for (int lineOffset = 1; index + lineOffset < lines.Length; lineOffset++) {
            var continuationLine = lines[index + lineOffset];
            if (string.IsNullOrWhiteSpace(continuationLine)) {
                return false;
            }

            var trimmedContinuation = continuationLine.TrimEnd();
            int closingBracket = FindReferenceLabelClosureOnContinuation(trimmedContinuation);
            if (closingBracket == -2) {
                return false;
            }

            if (closingBracket >= 0) {
                if (closingBracket + 1 >= trimmedContinuation.Length || trimmedContinuation[closingBracket + 1] != ':') {
                    return false;
                }

                labelBuilder.Append('\n');
                labelBuilder.Append(trimmedContinuation.Substring(0, closingBracket));
                label = NormalizeReferenceLabel(labelBuilder.ToString());
                int absoluteStartLine = state?.SourceLineOffset + index + 1 ?? index + 1;
                int absoluteClosingLine = state?.SourceLineOffset + index + lineOffset + 1 ?? index + lineOffset + 1;
                int labelStartLineIndex = index;
                int labelStartColumn = leading + 2;
                if (trimmed.Length == 1) {
                    FindMultilineReferenceLabelContentStart(lines, index, lineOffset, closingBracket, out labelStartLineIndex, out labelStartColumn);
                }

                int labelEndLineIndex = closingBracket > 0 ? index + lineOffset : index + lineOffset - 1;
                int labelEndColumn = closingBracket > 0
                    ? closingBracket
                    : GetReferenceLabelContinuationEndColumn(lines[labelEndLineIndex]);
                labelSpan = CreateSpan(
                    state,
                    state?.SourceLineOffset + labelStartLineIndex + 1 ?? labelStartLineIndex + 1,
                    labelStartColumn,
                    state?.SourceLineOffset + labelEndLineIndex + 1 ?? labelEndLineIndex + 1,
                    labelEndColumn);
                openingMarkerSpan = CreateSpan(state, absoluteStartLine, leading + 1, absoluteStartLine, leading + 1);
                separatorMarkerSpan = CreateSpan(state, absoluteClosingLine, closingBracket + 1, absoluteClosingLine, closingBracket + 2);

                int restStart = closingBracket + 2;
                while (restStart < trimmedContinuation.Length && char.IsWhiteSpace(trimmedContinuation[restStart])) {
                    restStart++;
                }

                rest = trimmedContinuation.Substring(restStart);
                consumedLines = lineOffset + 1;
                restLineIndex = index + lineOffset;
                restStartColumnZeroBased = restStart;
                return !string.IsNullOrEmpty(label);
            }

            labelBuilder.Append('\n');
            labelBuilder.Append(trimmedContinuation);
        }

        return false;
    }

    private static void FindMultilineReferenceLabelContentStart(
        string[] lines,
        int definitionStartIndex,
        int closingLineOffset,
        int closingBracket,
        out int labelStartLineIndex,
        out int labelStartColumn) {
        labelStartLineIndex = definitionStartIndex;
        labelStartColumn = 1;

        for (int lineOffset = 1; lineOffset <= closingLineOffset; lineOffset++) {
            int lineIndex = definitionStartIndex + lineOffset;
            string line = lines[lineIndex] ?? string.Empty;
            int endExclusive = lineOffset == closingLineOffset ? closingBracket : line.Length;
            while (endExclusive > 0 && char.IsWhiteSpace(line[endExclusive - 1])) {
                endExclusive--;
            }

            if (endExclusive > 0) {
                labelStartLineIndex = lineIndex;
                labelStartColumn = 1;
                return;
            }
        }
    }

    private static int GetReferenceLabelContinuationEndColumn(string line) {
        if (string.IsNullOrEmpty(line)) {
            return 1;
        }

        int endExclusive = line.Length;
        while (endExclusive > 0 && char.IsWhiteSpace(line[endExclusive - 1])) {
            endExclusive--;
        }

        return Math.Max(1, endExclusive);
    }

    private static int FindReferenceLabelClosureOnContinuation(string text) {
        if (string.IsNullOrEmpty(text)) {
            return -1;
        }

        bool escaped = false;
        for (int i = 0; i < text.Length; i++) {
            char c = text[i];
            if (escaped) {
                escaped = false;
                continue;
            }

            if (c == '\\') {
                escaped = true;
                continue;
            }

            if (c == '[') {
                return -2;
            }

            if (c == ']') {
                return i;
            }
        }

        return -1;
    }

    private static bool TryParseReferenceDestinationContinuation(
        string[] lines,
        int index,
        out string destinationLine,
        out int lineOffset,
        out int leadingWhitespace) {
        destinationLine = string.Empty;
        lineOffset = 0;
        leadingWhitespace = 0;
        if (lines == null || index < 0 || index >= lines.Length) {
            return false;
        }

        var line = lines[index];
        if (string.IsNullOrWhiteSpace(line)) {
            return false;
        }

        while (leadingWhitespace < line.Length && char.IsWhiteSpace(line[leadingWhitespace])) {
            leadingWhitespace++;
        }

        if (leadingWhitespace >= line.Length) {
            return false;
        }

        destinationLine = line.Substring(leadingWhitespace).TrimEnd();
        if (destinationLine.Length == 0) {
            return false;
        }

        lineOffset = 1;
        return true;
    }

    private static bool TryParseReferenceDestinationAndTitle(
        string rest,
        MarkdownReaderState? state,
        int lineIndex,
        int contentStartColumnZeroBased,
        out string url,
        out string? title,
        out MarkdownSourceSpan? urlSpan,
        out MarkdownSourceSpan? titleSpan) {
        url = string.Empty;
        title = null;
        urlSpan = null;
        titleSpan = null;

        if (string.IsNullOrEmpty(rest)) {
            return false;
        }

        if (TrySplitUrlAndOptionalTitle(
            rest,
            out url,
            out title,
            out int urlInnerStart,
            out int urlInnerLength,
            out int? titleInnerStart,
            out int? titleInnerLength)) {
            urlSpan = CreateSpan(
                state,
                state?.SourceLineOffset + lineIndex + 1 ?? lineIndex + 1,
                contentStartColumnZeroBased + urlInnerStart + 1,
                state?.SourceLineOffset + lineIndex + 1 ?? lineIndex + 1,
                contentStartColumnZeroBased + urlInnerStart + urlInnerLength);
            if (titleInnerStart.HasValue && titleInnerLength.HasValue) {
                titleSpan = CreateSpan(
                    state,
                    state?.SourceLineOffset + lineIndex + 1 ?? lineIndex + 1,
                    contentStartColumnZeroBased + titleInnerStart.Value + 1,
                    state?.SourceLineOffset + lineIndex + 1 ?? lineIndex + 1,
                    contentStartColumnZeroBased + titleInnerStart.Value + titleInnerLength.Value);
            }

            return true;
        }

        if (StartsWithAngleLinkDestination(rest) || IndexOfWhitespace(rest) >= 0) {
            return false;
        }

        url = DecodeLinkDestinationOrTitle(rest);
        title = null;
        urlSpan = CreateSpan(
            state,
            state?.SourceLineOffset + lineIndex + 1 ?? lineIndex + 1,
            contentStartColumnZeroBased + 1,
            state?.SourceLineOffset + lineIndex + 1 ?? lineIndex + 1,
            contentStartColumnZeroBased + rest.Length);
        return true;
    }

    private static bool StartsWithAngleLinkDestination(string value) {
        if (string.IsNullOrEmpty(value)) {
            return false;
        }

        int index = 0;
        while (index < value.Length && IsLinkWhitespace(value[index])) {
            index++;
        }

        return index < value.Length && value[index] == '<';
    }
}
