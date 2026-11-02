using System.Text;

namespace OfficeIMO.Markdown;

internal static partial class MarkdownReader {
    private static InlineSequence ParseJoinedParagraphLines(List<string> lines, MarkdownLoadOptions options, MarkdownReaderState? state) {
        var text = JoinParagraphLines(lines, options, out var softBreaks);
        return ParseInlines(text, options, state, CreateSemanticLineBreakMap(text, softBreaks));
    }

    // Token-only maps transport parsing semantics without allocating source points or capturing spans.
    private static MarkdownInlineSourceMap? CreateSemanticLineBreakMap(string text, List<int>? softBreaks) {
        if (softBreaks == null || softBreaks.Count == 0) return null;
        var tokens = new string?[text.Length];
        foreach (int offset in softBreaks) tokens[offset] = "\n";
        return new MarkdownInlineSourceMap(Array.Empty<MarkdownSourcePoint?>(), null, tokens);
    }

    private static (string Text, MarkdownInlineSourceMap? SourceMap) JoinTrimmedParagraphLines(List<string> lines, MarkdownLoadOptions options) {
        var text = JoinParagraphLines(lines, options, out var softBreaks);
        var sourceMap = CreateSemanticLineBreakMap(text, softBreaks);
        var trimmed = text.Trim();
        int start = text.Length - text.TrimStart().Length;
        return (trimmed, sourceMap?.Slice(start, trimmed.Length));
    }

    private static string JoinParagraphLines(List<string> lines, MarkdownLoadOptions options, out List<int>? softBreaks,
        MarkdownReaderState? state = null, int absoluteLineOffset = 0) {
        softBreaks = null;
        if (lines.Count == 1) {
            return GetParagraphLineJoinInfo(
                lines[0] ?? string.Empty,
                absoluteLine: 1,
                startColumn: 1,
                options,
                sourceTextMap: null,
                hasFollowingLine: false).Text;
        }

        var preservedInlineLineBreaks = FindParagraphLineBreaksInsideMatchedInlinePreserveSpans(lines, options);
        var sb = new StringBuilder(GetJoinedParagraphCapacity(lines));
        bool previousUsesLineBreakSeparator = false;
        bool previousHardBreak = false;
        bool previousPreserveLineBreak = false;
        for (int i = 0; i < lines.Count; i++) {
            var raw = lines[i] ?? string.Empty;
            if (previousHardBreak) raw = raw.TrimStart(' ', '\t');
            var joinInfo = GetParagraphLineJoinInfo(
                raw,
                absoluteLine: 1,
                startColumn: 1,
                options,
                sourceTextMap: null,
                hasFollowingLine: i + 1 < lines.Count,
                preserveLineEndingInsideInlineSpan: i < preservedInlineLineBreaks.Length && preservedInlineLineBreaks[i],
                preserveSoftLineBreakBeforeHtml: options.InlineHtml && i + 1 < lines.Count &&
                    StartsWithRawInlineHtml(lines[i + 1]));

            if (i > 0) {
                bool lazyQuoteBreak = !previousHardBreak &&
                    (IsLazyQuoteContinuationLine(state, absoluteLineOffset + i - 1) ||
                     IsLazyQuoteContinuationLine(state, absoluteLineOffset + i));
                if (previousPreserveLineBreak || lazyQuoteBreak) (softBreaks ??= new List<int>()).Add(sb.Length);
                sb.Append(previousUsesLineBreakSeparator || lazyQuoteBreak ? "\n" : " ");
            }
            sb.Append(joinInfo.Text);
            previousUsesLineBreakSeparator = joinInfo.UsesLineBreakSeparator;
            previousHardBreak = joinInfo.HardBreak;
            previousPreserveLineBreak = joinInfo.PreserveLineBreak && !joinInfo.HardBreak;
        }
        return sb.ToString();
    }

    private static (string Text, MarkdownInlineSourceMap? SourceMap) JoinParagraphLinesWithSourceMap(
        List<string> lines,
        int absoluteLineOffset,
        MarkdownLoadOptions options,
        MarkdownReaderState? state,
        IReadOnlyList<int>? lineStartColumns = null) {
        if (state?.SourceTextMap == null) {
            var sourceText = JoinParagraphLines(lines, options, out var softBreaks, state, absoluteLineOffset);
            return (sourceText, CreateSemanticLineBreakMap(sourceText, softBreaks));
        }

        var mapCapacity = GetJoinedParagraphCapacity(lines);
        var textBuilder = new StringBuilder(mapCapacity);
        var points = new MarkdownSourcePoint?[mapCapacity];
        MarkdownSourceSpan?[]? tokenSpans = null;
        string?[]? tokenLiterals = null;
        var mapLength = 0;
        var preservedInlineLineBreaks = FindParagraphLineBreaksInsideMatchedInlinePreserveSpans(lines, options);
        ParagraphLineJoinInfo? previousJoinInfo = null;
        var previousAbsoluteLine = absoluteLineOffset + 1;
        var previousJoinColumn = 1;

        for (var i = 0; i < lines.Count; i++) {
            var absoluteLine = absoluteLineOffset + i + 1;
            var lineStartColumn = GetParagraphLineStartColumn(lineStartColumns, i);
            if (previousJoinInfo.HasValue) {
                var softLazyQuoteBreak = (IsLazyQuoteContinuationLine(state, absoluteLineOffset + i - 1) ||
                    IsLazyQuoteContinuationLine(state, absoluteLineOffset + i)) &&
                    !previousJoinInfo.Value.HardBreak;
                textBuilder.Append(previousJoinInfo.Value.UsesLineBreakSeparator || softLazyQuoteBreak ? '\n' : ' ');
                points[mapLength] = state.SourceTextMap.CreatePoint(previousAbsoluteLine, previousJoinColumn);
                var tokenSpan = previousJoinInfo.Value.HardBreak
                    ? previousJoinInfo.Value.HardBreakMarkerSpan
                    : softLazyQuoteBreak || previousJoinInfo.Value.PreserveLineBreak
                        ? CreateSpan(
                            state,
                            previousAbsoluteLine,
                            previousJoinColumn,
                            previousAbsoluteLine,
                            previousJoinColumn)
                        : null;
                var tokenLiteral = previousJoinInfo.Value.HardBreak
                    ? previousJoinInfo.Value.HardBreakMarker
                    : softLazyQuoteBreak || previousJoinInfo.Value.PreserveLineBreak
                        ? "\n"
                        : null;
                SetInlineSourceMapToken(
                    ref tokenSpans,
                    ref tokenLiterals,
                    mapCapacity,
                    mapLength,
                    tokenSpan,
                    tokenLiteral);
                mapLength++;
            }

            var raw = lines[i] ?? string.Empty;
            if (previousJoinInfo?.HardBreak == true) {
                var trimmedStart = raw.TrimStart(' ', '\t');
                lineStartColumn = AdvanceParagraphSourceColumn(lineStartColumn, raw, raw.Length - trimmedStart.Length);
                raw = trimmedStart;
            }
            var joinInfo = GetParagraphLineJoinInfo(
                raw,
                absoluteLine,
                lineStartColumn,
                options,
                state.SourceTextMap,
                hasFollowingLine: i + 1 < lines.Count,
                preserveLineEndingInsideInlineSpan: i < preservedInlineLineBreaks.Length && preservedInlineLineBreaks[i],
                preserveSoftLineBreakBeforeHtml: options.InlineHtml && i + 1 < lines.Count &&
                    StartsWithRawInlineHtml(lines[i + 1]));
            textBuilder.Append(joinInfo.Text);
            var sourceColumn = state.SourceTextMap.WriteSequentialPoints(
                points,
                ref mapLength,
                joinInfo.Text,
                absoluteLine,
                lineStartColumn);

            previousAbsoluteLine = absoluteLine;
            previousJoinColumn = Math.Max(lineStartColumn, sourceColumn - 1);
            previousJoinInfo = joinInfo;
        }

        var text = textBuilder.ToString();
        if (string.IsNullOrEmpty(text)) {
            return (text, null);
        }

        ResizeInlineSourceMapArrays(ref points, ref tokenSpans, ref tokenLiterals, mapLength);
        return (text, new MarkdownInlineSourceMap(points, tokenSpans, tokenLiterals));
    }

    private static int GetParagraphLineStartColumn(IReadOnlyList<int>? lineStartColumns, int lineIndex) {
        if (lineStartColumns == null
            || lineIndex < 0
            || lineIndex >= lineStartColumns.Count
            || lineStartColumns[lineIndex] <= 0) {
            return 1;
        }

        return lineStartColumns[lineIndex];
    }

    private static (string Text, MarkdownInlineSourceMap? SourceMap) JoinParagraphSourceLinesWithSourceMap(
        List<MarkdownSourceLineSlice> lines,
        MarkdownLoadOptions options,
        MarkdownReaderState? state) {
        if (lines == null || lines.Count == 0) {
            return (string.Empty, null);
        }

        if (lines.Count == 1 && state?.SourceTextMap != null) {
            var line = lines[0];
            var joinInfo = GetParagraphLineJoinInfo(
                line.Text,
                line.AbsoluteLine,
                line.StartColumn,
                options,
                state.SourceTextMap,
                hasFollowingLine: false);
            return (
                joinInfo.Text,
                string.IsNullOrEmpty(joinInfo.Text)
                    ? null
                    : new MarkdownInlineSourceMap(
                        state.SourceTextMap,
                        joinInfo.Text,
                        line.AbsoluteLine,
                        line.StartColumn));
        }

        var plainLines = new List<string>(lines.Count);
        for (int i = 0; i < lines.Count; i++) {
            plainLines.Add(lines[i].Text);
        }

        if (state?.SourceTextMap == null) {
            var sourceText = JoinParagraphSourceLines(lines, options, out var softBreaks);
            return (sourceText, CreateSemanticLineBreakMap(sourceText, softBreaks));
        }

        var mapCapacity = GetJoinedParagraphCapacity(lines);
        var textBuilder = new StringBuilder(mapCapacity);
        var points = new MarkdownSourcePoint?[mapCapacity];
        MarkdownSourceSpan?[]? tokenSpans = null;
        string?[]? tokenLiterals = null;
        var mapLength = 0;
        var preservedInlineLineBreaks = FindParagraphLineBreaksInsideMatchedInlinePreserveSpans(plainLines, options);
        MarkdownSourceLineSlice? previousLine = null;
        ParagraphLineJoinInfo? previousJoinInfo = null;

        for (var i = 0; i < lines.Count; i++) {
            var slice = lines[i];
            if (previousJoinInfo?.HardBreak == true) {
                var trimmedStart = slice.Text.TrimStart(' ', '\t');
                slice = new MarkdownSourceLineSlice(trimmedStart, slice.AbsoluteLine,
                    AdvanceParagraphSourceColumn(slice.StartColumn, slice.Text, slice.Text.Length - trimmedStart.Length),
                    slice.IsLazyQuoteContinuation, slice.IsQuoteContainerLine);
            }
            if (previousLine.HasValue && previousJoinInfo.HasValue) {
                var softLazyQuoteBreak = IsLazyQuoteSoftBreak(previousLine.Value, slice) &&
                    !previousJoinInfo.Value.HardBreak;
                textBuilder.Append(previousJoinInfo.Value.UsesLineBreakSeparator || softLazyQuoteBreak ? '\n' : ' ');
                var previousJoinColumn = GetEndColumn(previousLine.Value.StartColumn, previousJoinInfo.Value.Text);
                points[mapLength] = state.SourceTextMap.CreatePoint(previousLine.Value.AbsoluteLine, previousJoinColumn);
                var tokenSpan = previousJoinInfo.Value.HardBreak
                    ? previousJoinInfo.Value.HardBreakMarkerSpan
                    : softLazyQuoteBreak || previousJoinInfo.Value.PreserveLineBreak
                        ? CreateSpan(
                            state,
                            previousLine.Value.AbsoluteLine,
                            previousJoinColumn,
                            previousLine.Value.AbsoluteLine,
                            previousJoinColumn)
                        : null;
                var tokenLiteral = previousJoinInfo.Value.HardBreak
                    ? previousJoinInfo.Value.HardBreakMarker
                    : softLazyQuoteBreak || previousJoinInfo.Value.PreserveLineBreak
                        ? "\n"
                        : null;
                SetInlineSourceMapToken(
                    ref tokenSpans,
                    ref tokenLiterals,
                    mapCapacity,
                    mapLength,
                    tokenSpan,
                    tokenLiteral);
                mapLength++;
            }

            var joinInfo = GetParagraphLineJoinInfo(
                slice.Text,
                slice.AbsoluteLine,
                slice.StartColumn,
                options,
                state.SourceTextMap,
                hasFollowingLine: i + 1 < lines.Count,
                preserveLineEndingInsideInlineSpan: i < preservedInlineLineBreaks.Length && preservedInlineLineBreaks[i],
                preserveSoftLineBreakBeforeHtml: options.InlineHtml && i + 1 < lines.Count &&
                    StartsWithRawInlineHtml(lines[i + 1].Text));
            textBuilder.Append(joinInfo.Text);
            state.SourceTextMap.WriteSequentialPoints(
                points,
                ref mapLength,
                joinInfo.Text,
                slice.AbsoluteLine,
                slice.StartColumn);

            previousLine = slice;
            previousJoinInfo = joinInfo;
        }

        var text = textBuilder.ToString();
        if (string.IsNullOrEmpty(text)) {
            return (text, null);
        }

        ResizeInlineSourceMapArrays(ref points, ref tokenSpans, ref tokenLiterals, mapLength);
        return (text, new MarkdownInlineSourceMap(points, tokenSpans, tokenLiterals));
    }

    private static string JoinParagraphSourceLines(List<MarkdownSourceLineSlice> lines, MarkdownLoadOptions options, out List<int>? softBreaks) {
        softBreaks = null;
        if (lines == null || lines.Count == 0) {
            return string.Empty;
        }

        if (lines.Count == 1) {
            return GetParagraphLineJoinInfo(
                lines[0].Text,
                lines[0].AbsoluteLine,
                lines[0].StartColumn,
                options,
                sourceTextMap: null,
                hasFollowingLine: false).Text;
        }

        var plainLines = new List<string>(lines.Count);
        for (int i = 0; i < lines.Count; i++) {
            plainLines.Add(lines[i].Text);
        }

        var preservedInlineLineBreaks = FindParagraphLineBreaksInsideMatchedInlinePreserveSpans(plainLines, options);
        var sb = new StringBuilder(GetJoinedParagraphCapacity(lines));
        ParagraphLineJoinInfo? previousJoinInfo = null;
        MarkdownSourceLineSlice? previousLine = null;
        for (int i = 0; i < lines.Count; i++) {
            var slice = lines[i];
            if (previousJoinInfo?.HardBreak == true) {
                var trimmedStart = slice.Text.TrimStart(' ', '\t');
                slice = new MarkdownSourceLineSlice(trimmedStart, slice.AbsoluteLine,
                    AdvanceParagraphSourceColumn(slice.StartColumn, slice.Text, slice.Text.Length - trimmedStart.Length),
                    slice.IsLazyQuoteContinuation, slice.IsQuoteContainerLine);
            }
            if (previousLine.HasValue && previousJoinInfo.HasValue) {
                if (!previousJoinInfo.Value.HardBreak && (previousJoinInfo.Value.PreserveLineBreak || IsLazyQuoteSoftBreak(previousLine.Value, slice))) {
                    (softBreaks ??= new List<int>()).Add(sb.Length);
                }
                sb.Append(previousJoinInfo.Value.UsesLineBreakSeparator ||
                          (IsLazyQuoteSoftBreak(previousLine.Value, slice) && !previousJoinInfo.Value.HardBreak)
                    ? '\n'
                    : ' ');
            }

            var joinInfo = GetParagraphLineJoinInfo(
                slice.Text,
                slice.AbsoluteLine,
                slice.StartColumn,
                options,
                sourceTextMap: null,
                hasFollowingLine: i + 1 < lines.Count,
                preserveLineEndingInsideInlineSpan: i < preservedInlineLineBreaks.Length && preservedInlineLineBreaks[i],
                preserveSoftLineBreakBeforeHtml: options.InlineHtml && i + 1 < lines.Count &&
                    StartsWithRawInlineHtml(lines[i + 1].Text));
            sb.Append(joinInfo.Text);
            previousLine = slice;
            previousJoinInfo = joinInfo;
        }

        return sb.ToString();
    }

    private static int GetJoinedParagraphCapacity(IReadOnlyList<string> lines) {
        var capacity = Math.Max(0, lines.Count - 1);
        for (var i = 0; i < lines.Count; i++) {
            capacity = checked(capacity + (lines[i]?.Length ?? 0));
        }

        return capacity;
    }

    private static int GetJoinedParagraphCapacity(IReadOnlyList<MarkdownSourceLineSlice> lines) {
        var capacity = Math.Max(0, lines.Count - 1);
        for (var i = 0; i < lines.Count; i++) {
            capacity = checked(capacity + lines[i].Text.Length);
        }

        return capacity;
    }

    private static void ResizeInlineSourceMapArrays(
        ref MarkdownSourcePoint?[] points,
        ref MarkdownSourceSpan?[]? tokenSpans,
        ref string?[]? tokenLiterals,
        int length) {
        if (length == points.Length) {
            return;
        }

        Array.Resize(ref points, length);
        if (tokenSpans != null) {
            Array.Resize(ref tokenSpans, length);
        }
        if (tokenLiterals != null) {
            Array.Resize(ref tokenLiterals, length);
        }
    }

    private static void SetInlineSourceMapToken(
        ref MarkdownSourceSpan?[]? tokenSpans,
        ref string?[]? tokenLiterals,
        int capacity,
        int index,
        MarkdownSourceSpan? tokenSpan,
        string? tokenLiteral) {
        if (tokenSpan.HasValue) {
            tokenSpans ??= new MarkdownSourceSpan?[capacity];
            tokenSpans[index] = tokenSpan;
        }
        if (tokenLiteral != null) {
            tokenLiterals ??= new string?[capacity];
            tokenLiterals[index] = tokenLiteral;
        }
    }

    private static bool IsLazyQuoteSoftBreak(MarkdownSourceLineSlice previousLine, MarkdownSourceLineSlice currentLine) =>
        previousLine.IsLazyQuoteContinuation || currentLine.IsLazyQuoteContinuation;

    private static bool IsLazyQuoteContinuationLine(MarkdownReaderState? state, int zeroBasedLineIndex) =>
        state?.LazyQuoteContinuationLines.Contains(zeroBasedLineIndex) == true;

    private static MarkdownInlineSourceMap? BuildInlineSourceMapForSingleLine(
        string text,
        int absoluteLine,
        int startColumn,
        MarkdownReaderState? state) {
        if (state?.SourceTextMap == null || string.IsNullOrEmpty(text)) {
            return null;
        }

        var points = new MarkdownSourcePoint?[text.Length];
        var destinationIndex = 0;
        state.SourceTextMap.WriteSequentialPoints(points, ref destinationIndex, text, absoluteLine, startColumn);

        return new MarkdownInlineSourceMap(points);
    }

    private static int GetEndColumn(int startColumn, string? text) {
        if (string.IsNullOrEmpty(text)) {
            return startColumn;
        }

        var sourceColumn = startColumn;
        for (var i = 0; i < text!.Length; i++) {
            sourceColumn = MarkdownSourceColumns.AdvanceColumn(sourceColumn, text[i]);
        }

        return Math.Max(startColumn, sourceColumn - 1);
    }

    private static string ConsumeTrailingBackslashHardBreak(string trimmed, MarkdownLoadOptions options, out bool hardBreak) {
        hardBreak = false;
        if (options == null || !options.BackslashHardBreaks) return trimmed ?? string.Empty;
        if (string.IsNullOrEmpty(trimmed)) return string.Empty;
        if (trimmed[trimmed.Length - 1] != '\\') return trimmed;
        hardBreak = true;
        return trimmed.Substring(0, trimmed.Length - 1);
    }

    private static ParagraphLineJoinInfo GetParagraphLineJoinInfo(
        string raw,
        int absoluteLine,
        int startColumn,
        MarkdownLoadOptions options,
        MarkdownSourceTextMap? sourceTextMap,
        bool hasFollowingLine,
        bool preserveLineEndingInsideInlineSpan = false,
        bool preserveSoftLineBreakBeforeHtml = false) {
        raw ??= string.Empty;
        if (preserveLineEndingInsideInlineSpan) {
            return new ParagraphLineJoinInfo(raw, hardBreak: false, hardBreakMarker: null, hardBreakMarkerSpan: null, preserveLineBreak: true);
        }

        bool spaceHardBreak = EndsWithTwoSpacesLine(raw);
        var trimmed = raw.TrimEnd();
        bool backslashHardBreak = false;
        var text = hasFollowingLine
            ? ConsumeTrailingBackslashHardBreak(trimmed, options, out backslashHardBreak)
            : trimmed;

        if (backslashHardBreak) {
            var markerColumn = AdvanceParagraphSourceColumn(startColumn, trimmed, Math.Max(0, trimmed.Length - 1));
            return new ParagraphLineJoinInfo(
                text,
                hardBreak: true,
                hardBreakMarker: "\\",
                hardBreakMarkerSpan: sourceTextMap?.CreateSpan(absoluteLine, markerColumn, absoluteLine, markerColumn));
        }

        if (spaceHardBreak) {
            var markerStartIndex = raw.TrimEnd(' ').Length;
            var markerLength = raw.Length - markerStartIndex;
            var markerStartColumn = AdvanceParagraphSourceColumn(startColumn, raw, markerStartIndex);
            var markerEndColumn = AdvanceParagraphSourceColumn(startColumn, raw, raw.Length) - 1;
            return new ParagraphLineJoinInfo(
                text,
                hardBreak: true,
                hardBreakMarker: markerLength > 0 ? raw.Substring(markerStartIndex, markerLength) : string.Empty,
                hardBreakMarkerSpan: sourceTextMap?.CreateSpan(absoluteLine, markerStartColumn, absoluteLine, markerEndColumn));
        }

        if (hasFollowingLine && options?.SoftLineBreaksAsHardLineBreaks == true) {
            return new ParagraphLineJoinInfo(text, hardBreak: true, hardBreakMarker: null, hardBreakMarkerSpan: null);
        }

        return new ParagraphLineJoinInfo(text, hardBreak: false, hardBreakMarker: null, hardBreakMarkerSpan: null,
            preserveLineBreak: preserveSoftLineBreakBeforeHtml);
    }

    private static int AdvanceParagraphSourceColumn(int startColumn, string? text, int endExclusive) {
        var column = Math.Max(1, startColumn);
        var value = text ?? string.Empty;
        if (value.Length == 0 || endExclusive <= 0) {
            return column;
        }

        int limit = Math.Min(value.Length, endExclusive);
        for (int i = 0; i < limit; i++) {
            column = MarkdownSourceColumns.AdvanceColumn(column, value[i]);
        }

        return column;
    }

}
