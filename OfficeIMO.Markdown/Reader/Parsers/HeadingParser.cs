namespace OfficeIMO.Markdown;

public static partial class MarkdownReader {
    internal sealed class HeadingParser : IMarkdownBlockParser {
        public bool TryParse(string[] lines, ref int i, MarkdownReaderOptions options, MarkdownDoc doc, MarkdownReaderState state) {
            if (!options.Headings) return false;
            if (!TryGetAtxHeadingContentRange(lines[i], out int level, out int contentStart, out int contentEnd, out string text, out int closingMarkerStart, out int closingMarkerEnd)) return false;
            int effectiveContentEnd = contentEnd;
            int effectiveClosingMarkerStart = closingMarkerStart;
            int effectiveClosingMarkerEnd = closingMarkerEnd;
            MarkdownAttributeSet parsedAttributes = MarkdownAttributeSet.Empty;
            MarkdownSourceSpan? attributeSpan = null;
            string? attributeSourceText = null;
            if (ShouldParseHeadingGenericAttributes(options, state)
                && MarkdownGenericAttributeParser.TryConsumeTrailingAttributeBlock(text, out var headingText, out parsedAttributes, out var attributeStart, out var attributeEnd, requireLeadingWhitespace: true)) {
                text = headingText;
                effectiveContentEnd = contentStart + attributeStart;
                if (effectiveClosingMarkerStart < 0 &&
                    TryTrimAtxClosingMarkerBeforeAttribute(headingText, contentStart, out var textWithoutClosingMarker, out var detectedClosingMarkerStart, out var detectedClosingMarkerEnd, out var textEndBeforeMarker)) {
                    text = textWithoutClosingMarker;
                    effectiveContentEnd = textEndBeforeMarker;
                    effectiveClosingMarkerStart = detectedClosingMarkerStart;
                    effectiveClosingMarkerEnd = detectedClosingMarkerEnd;
                }

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
            if (ShouldSuppressAutoIdentifierForLiteralHeadingGenericAttribute(text, options, state)) {
                heading.SuppressAutomaticIdentifier();
            }
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
            if (effectiveClosingMarkerStart >= 0 && effectiveClosingMarkerEnd > effectiveClosingMarkerStart) {
                heading.SetClosingMarkerSourceInfo(
                    0,
                    effectiveClosingMarkerStart + 1,
                    effectiveClosingMarkerEnd,
                    CreateSpan(state, absoluteLineNumber, effectiveClosingMarkerStart + 1, absoluteLineNumber, effectiveClosingMarkerEnd));
            }
            doc.Add(heading);
            i++; return true;
        }

        private static bool TryTrimAtxClosingMarkerBeforeAttribute(
            string headingText,
            int contentStart,
            out string textWithoutClosingMarker,
            out int closingMarkerStart,
            out int closingMarkerEnd,
            out int textEndBeforeMarker) {
            textWithoutClosingMarker = headingText;
            closingMarkerStart = -1;
            closingMarkerEnd = -1;
            textEndBeforeMarker = contentStart + headingText.Length;

            if (string.IsNullOrEmpty(headingText)) {
                return false;
            }

            var end = headingText.Length;
            while (end > 0 && char.IsWhiteSpace(headingText[end - 1])) {
                end--;
            }

            var markerStart = end;
            while (markerStart > 0 && headingText[markerStart - 1] == '#') {
                markerStart--;
            }

            if (markerStart == end) {
                return false;
            }

            var beforeMarker = markerStart - 1;
            if (beforeMarker >= 0 && !char.IsWhiteSpace(headingText[beforeMarker])) {
                return false;
            }

            var textEnd = Math.Max(0, beforeMarker);
            while (textEnd > 0 && char.IsWhiteSpace(headingText[textEnd - 1])) {
                textEnd--;
            }

            textWithoutClosingMarker = headingText.Substring(0, textEnd);
            closingMarkerStart = contentStart + markerStart;
            closingMarkerEnd = contentStart + end;
            textEndBeforeMarker = contentStart + textEnd;
            return true;
        }
    }
}
