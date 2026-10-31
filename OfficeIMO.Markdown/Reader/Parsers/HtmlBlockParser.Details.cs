using System.Collections.Generic;

namespace OfficeIMO.Markdown;

internal static partial class MarkdownReader {
    internal sealed partial class HtmlBlockParser {
        private static bool TryParseDetails(string htmlContent, int startLineIndex, MarkdownLoadOptions options, MarkdownReaderState state, out DetailsBlock? block) {
            block = null;
            if (string.IsNullOrWhiteSpace(htmlContent)) return false;

            int tagEnd = htmlContent.IndexOf('>');
            if (tagEnd < 0) return false;

            int openingStart = htmlContent.IndexOf('<');
            if (openingStart < 0 || openingStart > tagEnd) return false;

            string openTag = htmlContent.Substring(0, tagEnd + 1).Trim();
            if (!openTag.StartsWith("<details", StringComparison.OrdinalIgnoreCase)) return false;

            bool isOpen = openTag.IndexOf(" open", StringComparison.OrdinalIgnoreCase) >= 0;

            int closeIdx = htmlContent.LastIndexOf("</details>", StringComparison.OrdinalIgnoreCase);
            if (closeIdx < 0 || closeIdx <= tagEnd) return false;
            int trailingStart = closeIdx + "</details>".Length;
            if (trailingStart < htmlContent.Length && !string.IsNullOrWhiteSpace(htmlContent.Substring(trailingStart))) {
                return false;
            }

            string inner = htmlContent.Substring(tagEnd + 1, closeIdx - tagEnd - 1);

            SummaryBlock? summary = null;
            int bodyStart = 0;
            int summaryStart = inner.IndexOf("<summary", StringComparison.OrdinalIgnoreCase);
            if (summaryStart >= 0) {
                int summaryTagEnd = inner.IndexOf('>', summaryStart);
                if (summaryTagEnd < 0) return false;
                int summaryClose = inner.IndexOf("</summary>", summaryTagEnd + 1, StringComparison.OrdinalIgnoreCase);
                if (summaryClose < 0) return false;

                int summaryStartInHtml = tagEnd + 1 + summaryStart;
                int summaryTagEndInHtml = tagEnd + 1 + summaryTagEnd;
                int summaryCloseInHtml = tagEnd + 1 + summaryClose;
                int summaryCloseEndInHtml = summaryCloseInHtml + "</summary>".Length - 1;
                int summaryTextStartInHtml = tagEnd + 1 + summaryTagEnd + 1;
                int summaryTextEndInHtml = summaryCloseInHtml - 1;
                string summarySourceText = inner.Substring(summaryTagEnd + 1, summaryClose - summaryTagEnd - 1);
                int summaryTrimStart = GetTrimStartLength(summarySourceText);
                int summaryTrimEndExclusive = GetTrimEndExclusive(summarySourceText);
                string summaryInner = summaryTrimStart < summaryTrimEndExclusive
                    ? summarySourceText.Substring(summaryTrimStart, summaryTrimEndExclusive - summaryTrimStart)
                    : string.Empty;
                var decoded = System.Net.WebUtility.HtmlDecode(summaryInner);
                var summarySourceMap = BuildDetailsSummaryInlineSourceMap(
                    summarySourceText,
                    summaryTrimStart,
                    summaryTrimEndExclusive,
                    decoded,
                    htmlContent,
                    summaryTextStartInHtml,
                    startLineIndex,
                    state);
                var inlines = ParseInlines(decoded, options, state, summarySourceMap);
                summary = new SummaryBlock(inlines) {
                    SyntaxSpan = CreateDetailsSourceSpan(state, htmlContent, startLineIndex, summaryStartInHtml, summaryCloseEndInHtml),
                    OpeningTag = inner.Substring(summaryStart, summaryTagEnd - summaryStart + 1),
                    ClosingTag = inner.Substring(summaryClose, "</summary>".Length),
                    SourceText = summarySourceText,
                    OpeningTagSourceSpan = CreateDetailsSourceSpan(state, htmlContent, startLineIndex, summaryStartInHtml, summaryTagEndInHtml),
                    ClosingTagSourceSpan = CreateDetailsSourceSpan(state, htmlContent, startLineIndex, summaryCloseInHtml, summaryCloseEndInHtml)
                };

                if (summaryTextStartInHtml <= summaryTextEndInHtml) {
                    summary.TextSourceSpan = CreateDetailsSourceSpan(state, htmlContent, startLineIndex, summaryTextStartInHtml, summaryTextEndInHtml);
                }

                bodyStart = summaryClose + "</summary>".Length;
            }

            string body = inner.Substring(bodyStart);
            int bodyStartInHtml = tagEnd + 1 + bodyStart;
            var bodySourceLines = CreateDetailsBodySourceLines(htmlContent, body, bodyStartInHtml, startLineIndex, state);
            var (childBlocks, syntaxChildren) = ParseNestedMarkdownBlocks(bodySourceLines, options, state);
            block = new DetailsBlock(summary, childBlocks, isOpen) {
                InsertBlankLineAfterSummary = body.StartsWith("\n\n", StringComparison.Ordinal),
                InsertBlankLineBeforeClosing = body.EndsWith("\n\n", StringComparison.Ordinal),
                OpeningTag = htmlContent.Substring(openingStart, tagEnd - openingStart + 1),
                ClosingTag = htmlContent.Substring(closeIdx, "</details>".Length),
                OpeningTagSourceSpan = CreateDetailsSourceSpan(state, htmlContent, startLineIndex, openingStart, tagEnd),
                ClosingTagSourceSpan = CreateDetailsSourceSpan(state, htmlContent, startLineIndex, closeIdx, closeIdx + "</details>".Length - 1)
            };
            block.SyntaxChildren = syntaxChildren;
            return true;
        }

        private static int GetTrimStartLength(string text) {
            if (string.IsNullOrEmpty(text)) {
                return 0;
            }

            var index = 0;
            while (index < text.Length && char.IsWhiteSpace(text[index])) {
                index++;
            }

            return index;
        }

        private static int GetTrimEndExclusive(string text) {
            if (string.IsNullOrEmpty(text)) {
                return 0;
            }

            var index = text.Length;
            while (index > 0 && char.IsWhiteSpace(text[index - 1])) {
                index--;
            }

            return index;
        }

        private static MarkdownInlineSourceMap? BuildDetailsSummaryInlineSourceMap(
            string summarySourceText,
            int trimStart,
            int trimEndExclusive,
            string decodedText,
            string htmlContent,
            int summaryTextStartInHtml,
            int startLineIndex,
            MarkdownReaderState state) {
            if (state.SourceTextMap == null || string.IsNullOrEmpty(decodedText)) {
                return null;
            }

            var points = new List<MarkdownSourcePoint?>(decodedText.Length);
            var tokenSpans = new List<MarkdownSourceSpan?>(decodedText.Length);
            var tokenLiterals = new List<string?>(decodedText.Length);
            for (var sourceIndex = trimStart; sourceIndex < trimEndExclusive;) {
                var ch = summarySourceText[sourceIndex];
                if (ch == '&') {
                    var semicolon = summarySourceText.IndexOf(';', sourceIndex + 1);
                    if (semicolon >= 0 && semicolon < trimEndExclusive) {
                        var entityLiteral = summarySourceText.Substring(sourceIndex, semicolon - sourceIndex + 1);
                        var decodedEntity = System.Net.WebUtility.HtmlDecode(entityLiteral);
                        if (!string.Equals(decodedEntity, entityLiteral, StringComparison.Ordinal) && !string.IsNullOrEmpty(decodedEntity)) {
                            var span = CreateDetailsSourceSpan(
                                state,
                                htmlContent,
                                startLineIndex,
                                summaryTextStartInHtml + sourceIndex,
                                summaryTextStartInHtml + semicolon);
                            for (var i = 0; i < decodedEntity.Length; i++) {
                                points.Add(CreateDetailsSourcePoint(
                                    state,
                                    htmlContent,
                                    startLineIndex,
                                    summaryTextStartInHtml + sourceIndex));
                                tokenSpans.Add(i == 0 ? span : null);
                                tokenLiterals.Add(i == 0 ? entityLiteral : null);
                            }

                            sourceIndex = semicolon + 1;
                            continue;
                        }
                    }
                }

                points.Add(CreateDetailsSourcePoint(
                    state,
                    htmlContent,
                    startLineIndex,
                    summaryTextStartInHtml + sourceIndex));
                tokenSpans.Add(null);
                tokenLiterals.Add(null);
                sourceIndex++;
            }

            return points.Count == decodedText.Length
                ? new MarkdownInlineSourceMap(points.ToArray(), tokenSpans.ToArray(), tokenLiterals.ToArray())
                : null;
        }

        private static MarkdownSourcePoint CreateDetailsSourcePoint(
            MarkdownReaderState state,
            string htmlContent,
            int startLineIndex,
            int index) {
            GetDetailsSourcePosition(htmlContent, index, out int lineOffset, out int lineStart);
            int absoluteLine = state.SourceLineOffset + startLineIndex + lineOffset + 1;
            int column = AdvanceDetailsSourceColumn(htmlContent, lineStart, index);
            return state.SourceTextMap?.CreatePoint(absoluteLine, column)
                ?? new MarkdownSourcePoint(absoluteLine, column, 0);
        }

        private static IReadOnlyList<MarkdownSourceLineSlice> CreateDetailsBodySourceLines(
            string htmlContent,
            string body,
            int bodyStartInHtml,
            int startLineIndex,
            MarkdownReaderState state) {
            if (string.IsNullOrEmpty(body)) {
                return Array.Empty<MarkdownSourceLineSlice>();
            }

            var sourceLines = new List<MarkdownSourceLineSlice>();
            int lineStart = 0;
            while (lineStart <= body.Length) {
                int lineEnd = lineStart;
                while (lineEnd < body.Length && body[lineEnd] != '\r' && body[lineEnd] != '\n') {
                    lineEnd++;
                }

                string text = body.Substring(lineStart, lineEnd - lineStart);
                int sourceIndex = Math.Min(bodyStartInHtml + lineStart, Math.Max(0, htmlContent.Length - 1));
                GetDetailsSourcePosition(htmlContent, sourceIndex, out int lineOffset, out int detailLineStart);
                int absoluteLine = state.SourceLineOffset + startLineIndex + lineOffset + 1;
                int column = AdvanceDetailsSourceColumn(htmlContent, detailLineStart, sourceIndex);
                sourceLines.Add(new MarkdownSourceLineSlice(text, absoluteLine, column));

                if (lineEnd >= body.Length) {
                    break;
                }

                lineStart = lineEnd + 1;
                if (body[lineEnd] == '\r' && lineStart < body.Length && body[lineStart] == '\n') {
                    lineStart++;
                }
            }

            return sourceLines;
        }

        private static MarkdownSourceSpan? CreateDetailsSourceSpan(
            MarkdownReaderState state,
            string htmlContent,
            int startLineIndex,
            int startIndexInclusive,
            int endIndexInclusive) {
            GetDetailsSourcePosition(htmlContent, startIndexInclusive, out int startLineOffset, out int startLineStart);
            GetDetailsSourcePosition(htmlContent, endIndexInclusive, out int endLineOffset, out int endLineStart);

            int absoluteStartLine = state.SourceLineOffset + startLineIndex + startLineOffset + 1;
            int absoluteEndLine = state.SourceLineOffset + startLineIndex + endLineOffset + 1;
            int startColumn = AdvanceDetailsSourceColumn(htmlContent, startLineStart, startIndexInclusive);
            int endColumn = AdvanceDetailsSourceColumn(htmlContent, endLineStart, endIndexInclusive + 1) - 1;

            return CreateSpan(
                state,
                absoluteStartLine,
                startColumn,
                absoluteEndLine,
                endColumn);
        }

        private static void GetDetailsSourcePosition(string htmlContent, int index, out int lineOffset, out int lineStart) {
            int clampedIndex = Math.Max(0, Math.Min(index, Math.Max(0, htmlContent.Length - 1)));
            lineOffset = 0;
            lineStart = 0;

            for (int i = 0; i < clampedIndex; i++) {
                if (htmlContent[i] == '\n') {
                    lineOffset++;
                    lineStart = i + 1;
                }
            }
        }

        private static int AdvanceDetailsSourceColumn(string htmlContent, int startIndex, int endExclusive) {
            var column = 1;
            var boundedStart = Math.Max(0, Math.Min(startIndex, htmlContent.Length));
            var boundedEnd = Math.Max(boundedStart, Math.Min(endExclusive, htmlContent.Length));
            for (var i = boundedStart; i < boundedEnd; i++) {
                column = MarkdownSourceColumns.AdvanceColumn(column, htmlContent[i]);
            }

            return column;
        }

        private static int CountNewLines(string text, int start, int length) {
            if (string.IsNullOrEmpty(text) || length <= 0 || start >= text.Length) return 0;

            int end = Math.Min(text.Length, start + length);
            int count = 0;
            for (int i = start; i < end; i++) {
                if (text[i] == '\n') count++;
            }
            return count;
        }

    }
}
