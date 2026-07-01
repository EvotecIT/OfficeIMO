namespace OfficeIMO.Markdown;

public static partial class MarkdownReader {
    private static bool TryUpdateQuotedFencePreScanState(string line, ref bool inFence, ref char fenceChar, ref int fenceLen) {
        if (!TryStripQuotedPreScanLine(line, out var content, out _)) {
            if (inFence && !string.IsNullOrWhiteSpace(line)) {
                inFence = false;
            }

            return false;
        }

        if (!inFence) {
            if (IsCodeFenceOpen(content, out _, out fenceChar, out fenceLen)) {
                inFence = true;
                return true;
            }

            return false;
        }

        if (IsCodeFenceClose(content, fenceChar, fenceLen)) {
            inFence = false;
        }

        return true;
    }

    private static bool TryParseQuotedReferenceLinkDefinition(
        string[] lines,
        int index,
        MarkdownReaderOptions options,
        MarkdownReaderState state,
        out MarkdownReferenceLinkDefinition definition,
        out int consumedLines) {
        definition = null!;
        consumedLines = 0;

        if (lines == null || index < 0 || index >= lines.Length) {
            return false;
        }

        var strippedLines = new string[lines.Length];
        var contentStartColumns = new int[lines.Length];
        for (int i = index; i < lines.Length; i++) {
            if (TryStripQuotedPreScanLine(lines[i], out var content, out var contentStartColumn)) {
                strippedLines[i] = content;
                contentStartColumns[i] = contentStartColumn;
                continue;
            }

            strippedLines[i] = string.Empty;
        }

        if (contentStartColumns[index] == 0) {
            return false;
        }

        if (!TryParseReferenceLinkDefinition(
            strippedLines,
            index,
            options,
            state: null,
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

        for (int i = index; i < index + consumedLines && i < contentStartColumns.Length; i++) {
            if (!string.IsNullOrWhiteSpace(strippedLines[i]) && contentStartColumns[i] == 0) {
                return false;
            }
        }

        var resolved = ResolveUrl(url, options);
        if (resolved == null) {
            return false;
        }

        var sourceSpan = CreateLineSpan(
            state,
            state.SourceLineOffset + index + 1,
            state.SourceLineOffset + index + consumedLines);

        definition = new MarkdownReferenceLinkDefinition(
            label,
            resolved,
            title,
            sourceSpan,
            RemapQuotedReferenceSourceSpan(labelSpan, contentStartColumns, state),
            RemapQuotedReferenceSourceSpan(urlSpan, contentStartColumns, state),
            RemapQuotedReferenceSourceSpan(titleSpan, contentStartColumns, state),
            RemapQuotedReferenceSourceSpan(openingMarkerSpan, contentStartColumns, state),
            RemapQuotedReferenceSourceSpan(separatorMarkerSpan, contentStartColumns, state));
        return true;
    }

    private static bool TryStripQuotedPreScanLine(string line, out string content, out int contentStartColumn) {
        content = string.Empty;
        contentStartColumn = 0;

        if (string.IsNullOrEmpty(line)) {
            return false;
        }

        if (CountLeadingIndentColumns(line) > 3) {
            return false;
        }

        var trimmed = line.TrimStart();
        if (trimmed.Length == 0 || trimmed[0] != '>') {
            return false;
        }

        content = StripSingleQuoteMarker(line);
        contentStartColumn = GetQuoteContentStartColumn(line);
        return true;
    }

    private static MarkdownSourceSpan? RemapQuotedReferenceSourceSpan(
        MarkdownSourceSpan? span,
        int[] contentStartColumns,
        MarkdownReaderState state) {
        if (!span.HasValue) {
            return null;
        }

        var value = span.Value;
        int startIndex = value.StartLine - 1;
        int endIndex = value.EndLine - 1;
        if (startIndex < 0 || startIndex >= contentStartColumns.Length ||
            endIndex < 0 || endIndex >= contentStartColumns.Length ||
            contentStartColumns[startIndex] == 0 ||
            contentStartColumns[endIndex] == 0) {
            return null;
        }

        int startLine = state.SourceLineOffset + value.StartLine;
        int endLine = state.SourceLineOffset + value.EndLine;
        if (!value.StartColumn.HasValue || !value.EndColumn.HasValue) {
            return CreateLineSpan(state, startLine, endLine);
        }

        int startColumn = contentStartColumns[startIndex] + value.StartColumn.Value - 1;
        int endColumn = contentStartColumns[endIndex] + value.EndColumn.Value - 1;
        return CreateSpan(state, startLine, startColumn, endLine, endColumn);
    }
}
