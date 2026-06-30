namespace OfficeIMO.Markdown;

public static partial class MarkdownReader {
    /// <summary>
    /// Parses Setext headings:
    ///   Title
    ///   =====  (level 1)
    ///   Title
    ///   -----  (level 2)
    /// Requires one or more underline characters and no other content on the underline line.
    /// </summary>
    internal sealed class SetextHeadingParser : IMarkdownBlockParser {
        public bool TryParse(string[] lines, ref int i, MarkdownReaderOptions options, MarkdownDoc doc, MarkdownReaderState state) {
            if (!options.Headings) return false;
            if (i + 1 >= lines.Length) return false;
            var line = lines[i];
            var next = lines[i + 1];
            if (string.IsNullOrWhiteSpace(line) || string.IsNullOrWhiteSpace(next)) return false;
            if (IsSetextHeadingUnderlineSuppressed(state, i + 1)) return false;
            if (!TryGetSetextHeadingUnderlineLevel(next, out int level)) return false;
            var t = next.Trim();
            var headingText = line.Trim();
            var contentStart = line.IndexOf(headingText, StringComparison.Ordinal);
            var effectiveHeadingEnd = contentStart + headingText.Length;
            MarkdownAttributeSet parsedAttributes = MarkdownAttributeSet.Empty;
            MarkdownSourceSpan? attributeSpan = null;
            string? attributeSourceText = null;
            if (ShouldParseBlockGenericAttributes(options, state)
                && MarkdownGenericAttributeParser.TryConsumeTrailingAttributeBlock(headingText, out var textWithoutAttributeBlock, out parsedAttributes, out var attributeStart, out var attributeEnd, requireLeadingWhitespace: true)) {
                headingText = textWithoutAttributeBlock;
                effectiveHeadingEnd = contentStart + attributeStart;
                while (effectiveHeadingEnd > contentStart && char.IsWhiteSpace(line[effectiveHeadingEnd - 1])) {
                    effectiveHeadingEnd--;
                }

                var absoluteLineNumber = state.SourceLineOffset + i + 1;
                attributeSourceText = line.Substring(contentStart + attributeStart, attributeEnd - attributeStart + 1);
                attributeSpan = CreateSpan(
                    state,
                    absoluteLineNumber,
                    contentStart + attributeStart + 1,
                    absoluteLineNumber,
                    contentStart + attributeEnd + 1);
            }

            var sourceMap = BuildInlineSourceMapForSingleLine(headingText, state.SourceLineOffset + i + 1, contentStart + 1, state);
            var heading = new HeadingBlock(level, ParseInlines(headingText, options, state, sourceMap));
            heading.SetAttributes(parsedAttributes);
            MarkdownGenericAttributeSourceSpans.Set(heading, attributeSourceText, attributeSpan);
            var markerStartColumn = next.IndexOf(t, StringComparison.Ordinal) + 1;
            var markerEndColumn = markerStartColumn + t.Length - 1;
            var absoluteMarkerLine = state.SourceLineOffset + i + 2;
            heading.SetLevelSourceInfo(1, markerStartColumn, markerEndColumn);
            heading.SetSetextUnderlineMarkerSourceInfo(
                1,
                markerStartColumn,
                markerEndColumn,
                t,
                CreateSpan(state, absoluteMarkerLine, markerStartColumn, absoluteMarkerLine, markerEndColumn));
            if (headingText.Length > 0) {
                heading.SetTextSourceInfo(0, contentStart + 1, effectiveHeadingEnd);
            }
            doc.Add(heading);
            i += 2; // consume both lines
            return true;
        }
    }
}
