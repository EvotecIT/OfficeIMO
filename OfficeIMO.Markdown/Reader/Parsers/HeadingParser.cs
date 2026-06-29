namespace OfficeIMO.Markdown;

public static partial class MarkdownReader {
    internal sealed class HeadingParser : IMarkdownBlockParser {
        public bool TryParse(string[] lines, ref int i, MarkdownReaderOptions options, MarkdownDoc doc, MarkdownReaderState state) {
            if (!options.Headings) return false;
            if (!TryGetAtxHeadingContentRange(lines[i], out int level, out int contentStart, out int contentEnd, out string text, out int closingMarkerStart, out int closingMarkerEnd)) return false;
            int effectiveContentEnd = contentEnd;
            MarkdownAttributeSet parsedAttributes = MarkdownAttributeSet.Empty;
            MarkdownSourceSpan? attributeSpan = null;
            string? attributeSourceText = null;
            if (ShouldParseBlockGenericAttributes(options, state)
                && MarkdownGenericAttributeParser.TryConsumeTrailingAttributeBlock(text, out var headingText, out parsedAttributes, out var attributeStart, out var attributeEnd, requireLeadingWhitespace: true)) {
                text = headingText;
                effectiveContentEnd = contentStart + attributeStart;
                while (effectiveContentEnd > contentStart && char.IsWhiteSpace(lines[i][effectiveContentEnd - 1])) {
                    effectiveContentEnd--;
                }

                var attributeLineNumber = state.SourceLineOffset + i + 1;
                attributeSourceText = lines[i].Substring(contentStart + attributeStart, attributeEnd - attributeStart + 1);
                attributeSpan = CreateSpan(
                    state,
                    attributeLineNumber,
                    contentStart + attributeStart + 1,
                    attributeLineNumber,
                    contentStart + attributeEnd + 1);
            }

            var sourceMap = BuildInlineSourceMapForSingleLine(text, state.SourceLineOffset + i + 1, contentStart + 1, state);
            var heading = new HeadingBlock(level, ParseInlines(text, options, state, sourceMap));
            heading.SetAttributes(parsedAttributes);
            MarkdownGenericAttributeSourceSpans.Set(heading, attributeSourceText, attributeSpan);
            var markerStartColumn = CountLeadingSpaces(lines[i]) + 1;
            var markerEndColumn = markerStartColumn + level - 1;
            var absoluteLineNumber = state.SourceLineOffset + i + 1;
            heading.SetLevelSourceInfo(0, markerStartColumn, markerStartColumn + level - 1);
            heading.SetOpeningMarkerSourceInfo(
                0,
                markerStartColumn,
                markerEndColumn,
                CreateSpan(state, absoluteLineNumber, markerStartColumn, absoluteLineNumber, markerEndColumn));
            if (effectiveContentEnd > contentStart) {
                heading.SetTextSourceInfo(0, contentStart + 1, effectiveContentEnd);
            }
            if (closingMarkerStart >= 0 && closingMarkerEnd > closingMarkerStart) {
                heading.SetClosingMarkerSourceInfo(
                    0,
                    closingMarkerStart + 1,
                    closingMarkerEnd,
                    CreateSpan(state, absoluteLineNumber, closingMarkerStart + 1, absoluteLineNumber, closingMarkerEnd));
            }
            doc.Add(heading);
            i++; return true;
        }
    }
}
