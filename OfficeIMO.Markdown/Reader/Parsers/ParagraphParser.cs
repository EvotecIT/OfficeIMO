namespace OfficeIMO.Markdown;

public static partial class MarkdownReader {
    internal sealed class ParagraphParser : IMarkdownBlockParser {
        public bool TryParse(string[] lines, ref int i, MarkdownReaderOptions options, MarkdownDoc doc, MarkdownReaderState state) {
            if (!options.Paragraphs) return false;
            // Paragraph begins when none of the other block starters match.
            if (IsAtxHeading(lines[i], out _, out _) ||
                IsCodeFenceOpen(lines[i], out _, out _, out _) ||
                StartsTable(lines, i, options, state) ||
                IsParagraphInterruptingThematicBreakLine(lines[i]) ||
                IsParagraphInterruptingUnorderedListLine(lines[i]) ||
                IsOrderedListLine(lines[i], options, out _, out _) ||
                (options.Callouts && IsCalloutHeader(lines[i], options, out _, out _)) ||
                IsQuoteStarter(lines[i]) ||
                HtmlBlockParser.IsParagraphInterruptingHtmlBlockStart(lines[i], options) ||
                (IsReferenceLinkDefinitionStarter(lines, i, options) && !ShouldTreatReferenceDefinitionAsGenericAttributeParagraph(lines, i, options, state)) ||
                IsAbbreviationDefinitionStarter(lines[i], options) ||
                IsFootnoteDefinitionStarter(lines[i], options) ||
                (options.StandaloneImageBlocks && IsImageLine(lines[i]))) return false;

            var sb = new StringBuilder();
            int j = i;
            bool prevHard = false;
            while (j < lines.Length && !string.IsNullOrWhiteSpace(lines[j]) &&
                   !IsAtxHeading(lines[j], out _, out _) &&
                   !IsCodeFenceOpen(lines[j], out _, out _, out _) &&
                   !StartsTable(lines, j, options, state) &&
                   !IsParagraphInterruptingThematicBreakLine(lines[j]) &&
                   !IsParagraphInterruptingUnorderedListLine(lines[j]) &&
                   !IsParagraphInterruptingOrderedListLine(lines[j]) &&
                   (!options.Callouts || !IsCalloutHeader(lines[j], options, out _, out _)) &&
                   !IsQuoteStarter(lines[j]) &&
                   !HtmlBlockParser.IsParagraphInterruptingHtmlBlockStart(lines[j], options) &&
                   !IsParagraphTerminatingReferenceLinkDefinition(lines, i, j, options, state) &&
                   !IsAbbreviationDefinitionStarter(lines[j], options) &&
                   !IsFootnoteDefinitionStarter(lines[j], options) &&
                   !(options.StandaloneImageBlocks && IsImageLine(lines[j]))) {
                var raw = lines[j];
                bool hard = EndsWithTwoSpaces(raw);
                var trimmed = raw.TrimEnd();
                trimmed = ConsumeTrailingBackslashHardBreak(trimmed, options, out bool slashHard);
                hard = hard || slashHard;
                if (j > i) sb.Append(prevHard ? "\n" : " ");
                sb.Append(trimmed);
                prevHard = hard;
                j++;
            }
            if (sb.Length == 0) return false;
            var paragraphLines = new List<string>(j - i);
            for (var lineIndex = i; lineIndex < j; lineIndex++) {
                paragraphLines.Add(lines[lineIndex]);
            }

            var underlineLineIndex = i + paragraphLines.Count - 1;
            if (!IsSetextHeadingUnderlineSuppressed(state, underlineLineIndex) &&
                TryParseSetextHeadingParagraphLines(paragraphLines, options, out int level, out string headingText)) {
                var contentLines = paragraphLines.GetRange(0, paragraphLines.Count - 1);
                MarkdownAttributeSet headingAttributes = MarkdownAttributeSet.Empty;
                MarkdownSourceSpan? headingAttributeSpan = null;
                string? headingAttributeSourceText = null;
                if (ShouldParseHeadingGenericAttributes(options, state) && contentLines.Count > 0) {
                    var lastContentLineIndex = contentLines.Count - 1;
                    if (MarkdownGenericAttributeParser.TryConsumeTrailingAttributeBlock(
                        contentLines[lastContentLineIndex],
                        out var lineWithoutAttributeBlock,
                        out headingAttributes,
                        out var attributeStart,
                        out var attributeEnd,
                        requireLeadingWhitespace: true)) {
                        var attributeLine = contentLines[lastContentLineIndex];
                        var absoluteAttributeLine = state.SourceLineOffset + i + lastContentLineIndex + 1;
                        headingAttributeSourceText = attributeLine.Substring(attributeStart, attributeEnd - attributeStart + 1);
                        headingAttributeSpan = CreateSpan(
                            state,
                            absoluteAttributeLine,
                            attributeStart + 1,
                            absoluteAttributeLine,
                            attributeEnd + 1);
                        contentLines[lastContentLineIndex] = lineWithoutAttributeBlock;
                        while (contentLines.Count > 0 && string.IsNullOrWhiteSpace(contentLines[contentLines.Count - 1])) {
                            contentLines.RemoveAt(contentLines.Count - 1);
                        }
                    }
                }

                var (headingInlineText, headingSourceMap) = JoinParagraphLinesWithSourceMap(contentLines, state.SourceLineOffset + i, options, state);
                var heading = new HeadingBlock(level, ParseInlines(headingInlineText, options, state, headingSourceMap));
                if (contentLines.Count > 0
                    && ShouldSuppressAutoIdentifierForLiteralHeadingGenericAttribute(contentLines, options, state)) {
                    heading.SuppressAutomaticIdentifier();
                }
                heading.SetAttributes(headingAttributes);
                MarkdownGenericAttributeSourceSpans.Set(heading, headingAttributeSourceText, headingAttributeSpan);
                var underline = paragraphLines[paragraphLines.Count - 1] ?? string.Empty;
                var trimmedUnderline = underline.Trim();
                var markerStartColumn = underline.IndexOf(trimmedUnderline, StringComparison.Ordinal) + 1;
                var markerEndColumn = markerStartColumn + trimmedUnderline.Length - 1;
                var markerLineOffset = paragraphLines.Count - 1;
                var absoluteMarkerLine = state.SourceLineOffset + i + paragraphLines.Count;
                heading.SetLevelSourceInfo(markerLineOffset, markerStartColumn, markerEndColumn);
                heading.SetSetextUnderlineMarkerSourceInfo(
                    markerLineOffset,
                    markerStartColumn,
                    markerEndColumn,
                    trimmedUnderline,
                    CreateSpan(state, absoluteMarkerLine, markerStartColumn, absoluteMarkerLine, markerEndColumn));
                if (contentLines.Count > 0) {
                    heading.SetTextSourceInfo(
                        0,
                        GetFirstNonWhitespaceColumn(contentLines[0]),
                        contentLines.Count - 1,
                        GetTrimmedEndColumn(contentLines[contentLines.Count - 1]));
                }
                doc.Add(heading);
                i = j;
                return true;
            }

            MarkdownAttributeSet paragraphAttributes = MarkdownAttributeSet.Empty;
            MarkdownSourceSpan? paragraphAttributeSpan = null;
            string? paragraphAttributeSourceText = null;
            string paragraphAttributeConsumedWhitespace = string.Empty;
            bool suppressGenericAttributeSeparator = false;
            bool suppressGenericAttributeMetadata = false;
            bool suppressInlineAutolinks = false;
            if (ShouldParseParagraphGenericAttributes(options, state, i) && paragraphLines.Count > 0) {
                var lastLineIndex = paragraphLines.Count - 1;
                if (IsStandaloneGenericAttributeContinuationLine(paragraphLines, lastLineIndex)) {
                    paragraphLines.RemoveAt(lastLineIndex);
                    RemoveTrailingHardBreakMarkerBeforeConsumedAttributeLine(paragraphLines, options);
                } else if (!IsStandaloneGenericAttributeBeforeBlockquote(paragraphLines, lines, j)
                    && TryConsumeParagraphTrailingGenericAttributes(
                    paragraphLines[lastLineIndex],
                    options,
                    state,
                    out var lineWithoutAttributeBlock,
                    out paragraphAttributes,
                    out var attributeStart,
                    out var attributeEnd,
                    out paragraphAttributeConsumedWhitespace,
                    out suppressGenericAttributeSeparator,
                    out suppressGenericAttributeMetadata,
                    out suppressInlineAutolinks)) {
                    var attributeLine = paragraphLines[lastLineIndex];
                    var absoluteAttributeLine = state.SourceLineOffset + i + lastLineIndex + 1;

                    if (!suppressGenericAttributeMetadata) {
                        paragraphAttributeSourceText = attributeLine.Substring(attributeStart, attributeEnd - attributeStart + 1);
                        paragraphAttributeSpan = CreateSpan(
                            state,
                            absoluteAttributeLine,
                            attributeStart + 1,
                            absoluteAttributeLine,
                            attributeEnd + 1);
                    }

                    paragraphLines[lastLineIndex] = lineWithoutAttributeBlock;
                    while (paragraphLines.Count > 0 && string.IsNullOrWhiteSpace(paragraphLines[paragraphLines.Count - 1])) {
                        paragraphLines.RemoveAt(paragraphLines.Count - 1);
                    }
                }
            }

            ConsumeLeadingSoftBreakGenericAttributeContinuationLines(paragraphLines, options, state, i, out var paragraphLineStartColumns);
            var (text, sourceMap) = JoinParagraphLinesWithSourceMap(paragraphLines, state.SourceLineOffset + i, options, state, paragraphLineStartColumns);
            var inlineOptions = suppressInlineAutolinks ? CloneOptionsWithoutInlineAutolinks(options) : options;
            var paragraph = new ParagraphBlock(ParseInlines(text, inlineOptions, state, sourceMap));
            paragraph.SetAttributes(paragraphAttributes);
            paragraph.GenericAttributeConsumedWhitespace = paragraphAttributeConsumedWhitespace;
            paragraph.GenericAttributeSuppressSeparator = suppressGenericAttributeSeparator;
            MarkdownGenericAttributeSourceSpans.Set(paragraph, paragraphAttributeSourceText, paragraphAttributeSpan);
            if (!IsQuoteContainerParagraph(state, i, j)) {
                PromoteNestedInlineGenericAttributesToParagraph(paragraph, options);
            }
            doc.Add(paragraph);
            i = j; return true;
        }

        private static bool ShouldSuppressAutoIdentifierForLiteralHeadingGenericAttribute(
            IReadOnlyList<string> contentLines,
            MarkdownReaderOptions options,
            MarkdownReaderState state) {
            if (contentLines == null || contentLines.Count == 0) {
                return false;
            }

            for (int lineIndex = 0; lineIndex < contentLines.Count; lineIndex++) {
                if (MarkdownReader.ShouldSuppressAutoIdentifierForLiteralHeadingGenericAttribute(contentLines[lineIndex], options, state)) {
                    return true;
                }
            }

            return false;
        }

        private static bool IsQuoteContainerParagraph(MarkdownReaderState state, int startIndex, int endExclusiveIndex) {
            if (state == null || state.QuoteContainerLines.Count == 0) {
                return false;
            }

            for (var lineIndex = startIndex; lineIndex < endExclusiveIndex; lineIndex++) {
                if (state.QuoteContainerLines.Contains(lineIndex)) {
                    return true;
                }
            }

            return false;
        }

        private static bool TryConsumeParagraphTrailingGenericAttributes(
            string line,
            MarkdownReaderOptions options,
            MarkdownReaderState state,
            out string lineWithoutAttributeBlock,
            out MarkdownAttributeSet attributes,
            out int attributeStart,
            out int attributeEnd,
            out string consumedWhitespace,
            out bool suppressGenericAttributeSeparator,
            out bool suppressGenericAttributeMetadata,
            out bool suppressInlineAutolinks) {
            consumedWhitespace = string.Empty;
            suppressGenericAttributeSeparator = false;
            suppressGenericAttributeMetadata = false;
            suppressInlineAutolinks = false;

            if (MarkdownGenericAttributeParser.TryConsumeTrailingAttributeBlock(
                line,
                out lineWithoutAttributeBlock,
                out attributes,
                out attributeStart,
                out attributeEnd,
                requireLeadingWhitespace: true)) {
                if (attributeStart >= lineWithoutAttributeBlock.Length) {
                    consumedWhitespace = line.Substring(
                        lineWithoutAttributeBlock.Length,
                        attributeStart - lineWithoutAttributeBlock.Length);
                }

                return true;
            }

            if (!MarkdownGenericAttributeParser.TryConsumeTrailingAttributeBlock(
                line,
                out lineWithoutAttributeBlock,
                out attributes,
                out attributeStart,
                out attributeEnd,
                requireLeadingWhitespace: false)) {
                return false;
            }

            if (IsNoSpaceBareAutolinkParagraphAttribute(line, lineWithoutAttributeBlock, attributeStart, options)) {
                suppressGenericAttributeSeparator = true;
                suppressInlineAutolinks = true;
                return true;
            }

            if (IsNoSpaceAbbreviationParagraphAttribute(line, lineWithoutAttributeBlock, attributeStart, options, state)) {
                suppressGenericAttributeSeparator = true;
                return true;
            }

            if (IsNoSpaceCharacterReferenceParagraphAttribute(line, lineWithoutAttributeBlock, attributeStart)) {
                suppressGenericAttributeSeparator = true;
                suppressGenericAttributeMetadata = true;
                attributes = MarkdownAttributeSet.Empty;
                return true;
            }

            if (IsNoSpacePlainTextParagraphAttribute(line, lineWithoutAttributeBlock, attributeStart)) {
                suppressGenericAttributeSeparator = true;
                return true;
            }

            lineWithoutAttributeBlock = line;
            attributes = MarkdownAttributeSet.Empty;
            attributeStart = -1;
            attributeEnd = -1;
            return false;
        }

        private static bool IsStandaloneGenericAttributeContinuationLine(
            IReadOnlyList<string> paragraphLines,
            int lineIndex) =>
            paragraphLines != null
            && paragraphLines.Count > 1
            && lineIndex >= 0
            && lineIndex < paragraphLines.Count
            && IsStandaloneGenericAttributeOnlyLine(paragraphLines[lineIndex]);

        private static void RemoveTrailingHardBreakMarkerBeforeConsumedAttributeLine(
            List<string> paragraphLines,
            MarkdownReaderOptions options) {
            if (paragraphLines == null || paragraphLines.Count == 0) {
                return;
            }

            var previousIndex = paragraphLines.Count - 1;
            var previous = paragraphLines[previousIndex] ?? string.Empty;
            var trimmed = previous.TrimEnd();
            if (options?.BackslashHardBreaks == true && trimmed.EndsWith("\\", StringComparison.Ordinal)) {
                paragraphLines[previousIndex] = trimmed.Substring(0, trimmed.Length - 1);
                return;
            }

            if (EndsWithTwoSpaces(previous)) {
                paragraphLines[previousIndex] = trimmed;
            }
        }

        private static bool IsStandaloneGenericAttributeBeforeBlockquote(
            IReadOnlyList<string> paragraphLines,
            string[] lines,
            int nextLineIndex) {
            if (paragraphLines == null
                || paragraphLines.Count != 1
                || lines == null
                || nextLineIndex < 0
                || nextLineIndex >= lines.Length
                || !IsStandaloneGenericAttributeOnlyLine(paragraphLines[0])) {
                return false;
            }

            for (var index = nextLineIndex; index < lines.Length; index++) {
                if (string.IsNullOrWhiteSpace(lines[index])) {
                    continue;
                }

                return IsQuoteStarter(lines[index]);
            }

            return false;
        }

        private static bool IsStandaloneGenericAttributeOnlyLine(string? line) {
            if (string.IsNullOrWhiteSpace(line)) {
                return false;
            }

            var leading = CountLeadingSpaces(line!);
            var content = line!.Substring(leading).TrimEnd();
            return MarkdownGenericAttributeParser.TryConsumeLeadingAttributeBlock(
                    content,
                    out var remaining,
                    out _,
                    out var consumedLength)
                && consumedLength == content.Length
                && string.IsNullOrWhiteSpace(remaining);
        }

        private static void ConsumeLeadingSoftBreakGenericAttributeContinuationLines(
            List<string> paragraphLines,
            MarkdownReaderOptions options,
            MarkdownReaderState state,
            int paragraphStartLineIndex,
            out IReadOnlyList<int>? lineStartColumns) {
            lineStartColumns = null;
            if (!ShouldParseParagraphGenericAttributes(options, state, paragraphStartLineIndex)
                || paragraphLines == null
                || paragraphLines.Count < 2) {
                return;
            }

            List<int>? columns = null;
            for (var lineIndex = 1; lineIndex < paragraphLines.Count; lineIndex++) {
                var line = paragraphLines[lineIndex] ?? string.Empty;
                var leading = CountLeadingSpaces(line);
                var content = line.Substring(leading);
                if (!MarkdownGenericAttributeParser.TryConsumeLeadingAttributeBlock(
                        content,
                        out var remaining,
                        out var attributes,
                        out var consumedLength)
                    || attributes.IsEmpty
                    || string.IsNullOrWhiteSpace(remaining)) {
                    continue;
                }

                var remainingStart = 0;
                while (remainingStart < remaining.Length && char.IsWhiteSpace(remaining[remainingStart])) {
                    remainingStart++;
                }

                if (remainingStart >= remaining.Length) {
                    continue;
                }

                columns ??= CreateDefaultLineStartColumns(paragraphLines.Count);
                columns[lineIndex] = leading + consumedLength + remainingStart + 1;
                paragraphLines[lineIndex] = remaining.Substring(remainingStart);
            }

            lineStartColumns = columns;
        }

        private static List<int> CreateDefaultLineStartColumns(int count) {
            var columns = new List<int>(count);
            for (var i = 0; i < count; i++) {
                columns.Add(1);
            }

            return columns;
        }

        private static bool IsNoSpaceBareAutolinkParagraphAttribute(
            string line,
            string lineWithoutAttributeBlock,
            int attributeStart,
            MarkdownReaderOptions options) {
            if (string.IsNullOrEmpty(line)
                || string.IsNullOrWhiteSpace(lineWithoutAttributeBlock)
                || attributeStart <= 0
                || attributeStart > line.Length
                || char.IsWhiteSpace(line[attributeStart - 1])) {
                return false;
            }

            var candidate = lineWithoutAttributeBlock.Trim();
            if (candidate.Length == 0 || candidate.IndexOfAny(new[] { ' ', '\t', '\r', '\n' }) >= 0) {
                return false;
            }

            if (options.AutolinkUrls && StartsWithHttp(candidate, 0, options, out int httpEnd) && httpEnd == candidate.Length) {
                return true;
            }

            if (options.AutolinkWwwUrls && StartsWithWww(candidate, 0, options, out int wwwEnd) && wwwEnd == candidate.Length) {
                return true;
            }

            return options.AutolinkBareSchemeUrls
                && TryConsumeBareSchemeAutolink(candidate, 0, options, out int schemeEnd, out _, out _)
                && schemeEnd == candidate.Length;
        }

        private static bool IsNoSpaceAbbreviationParagraphAttribute(
            string line,
            string lineWithoutAttributeBlock,
            int attributeStart,
            MarkdownReaderOptions options,
            MarkdownReaderState state) {
            if (options?.Abbreviations != true
                || state == null
                || state.Abbreviations.Count == 0
                || string.IsNullOrEmpty(line)
                || string.IsNullOrEmpty(lineWithoutAttributeBlock)
                || attributeStart <= 0
                || attributeStart > line.Length
                || char.IsWhiteSpace(line[attributeStart - 1])) {
                return false;
            }

            for (int position = 0; position < lineWithoutAttributeBlock.Length; position++) {
                if (TryConsumeAbbreviation(lineWithoutAttributeBlock, position, state, out var definition)
                    && position + definition.Label.Length == lineWithoutAttributeBlock.Length) {
                    return true;
                }
            }

            return false;
        }

        private static bool IsNoSpacePlainTextParagraphAttribute(
            string line,
            string lineWithoutAttributeBlock,
            int attributeStart) {
            if (string.IsNullOrEmpty(line)
                || string.IsNullOrWhiteSpace(lineWithoutAttributeBlock)
                || attributeStart <= 0
                || attributeStart > line.Length
                || char.IsWhiteSpace(line[attributeStart - 1])) {
                return false;
            }

            var candidate = lineWithoutAttributeBlock.TrimEnd();
            if (EndsWithEscapedCharacterReference(candidate)) {
                return true;
            }

            if (EndsWithCharacterReference(candidate)) {
                return false;
            }

            return candidate.Length > 0
                && !EndsWithNoSpaceInlineGenericAttributeTarget(candidate);
        }

        private static bool IsNoSpaceCharacterReferenceParagraphAttribute(
            string line,
            string lineWithoutAttributeBlock,
            int attributeStart) {
            if (string.IsNullOrEmpty(line)
                || string.IsNullOrWhiteSpace(lineWithoutAttributeBlock)
                || attributeStart <= 0
                || attributeStart > line.Length
                || char.IsWhiteSpace(line[attributeStart - 1])) {
                return false;
            }

            var candidate = lineWithoutAttributeBlock.TrimEnd();
            return !EndsWithEscapedCharacterReference(candidate)
                && EndsWithCharacterReference(candidate);
        }

        private static bool EndsWithCharacterReference(string candidate) {
            if (string.IsNullOrEmpty(candidate) || candidate[candidate.Length - 1] != ';') {
                return false;
            }

            var ampersand = candidate.LastIndexOf('&');
            return ampersand >= 0
                && CommonMarkCharacterReference.TryDecode(candidate, ampersand, out int consumed, out _)
                && ampersand + consumed == candidate.Length;
        }

        private static bool EndsWithEscapedCharacterReference(string candidate) {
            if (string.IsNullOrEmpty(candidate) || candidate[candidate.Length - 1] != ';') {
                return false;
            }

            var ampersand = candidate.LastIndexOf('&');
            return ampersand > 0
                && candidate[ampersand - 1] == '\\'
                && !IsEscapedFinalCharacter(candidate.Substring(0, ampersand))
                && CommonMarkCharacterReference.TryDecode(candidate, ampersand, out int consumed, out _)
                && ampersand + consumed == candidate.Length;
        }

        private static bool EndsWithNoSpaceInlineGenericAttributeTarget(string candidate) {
            var last = candidate[candidate.Length - 1];
            if (IsEscapedFinalCharacter(candidate)) {
                return false;
            }

            if (last == ')'
                || last == ']'
                || last == '*'
                || last == '_'
                || last == '>') {
                return true;
            }

            if (last == '`') {
                return EndsWithClosingCodeSpan(candidate);
            }

            return last switch {
                '~' => EndsWithDelimitedRun(candidate, "~~") || HasEarlierDelimiter(candidate, "~", 1),
                '^' => HasEarlierDelimiter(candidate, "^", 1),
                '+' => EndsWithDelimitedRun(candidate, "++"),
                '=' => EndsWithDelimitedRun(candidate, "=="),
                _ => false
            };
        }

        private static bool IsEscapedFinalCharacter(string candidate) {
            if (string.IsNullOrEmpty(candidate) || candidate.Length < 2 || !IsBackslashEscapable(candidate[candidate.Length - 1])) {
                return false;
            }

            var slashCount = 0;
            for (var index = candidate.Length - 2; index >= 0 && candidate[index] == '\\'; index--) {
                slashCount++;
            }

            return slashCount % 2 == 1;
        }

        private static bool EndsWithClosingCodeSpan(string candidate) {
            var closingStart = candidate.Length - 1;
            while (closingStart > 0 && candidate[closingStart - 1] == '`') {
                closingStart--;
            }

            if (closingStart == 0) {
                return false;
            }

            var delimiter = candidate.Substring(closingStart);
            return candidate.LastIndexOf(delimiter, closingStart - 1, StringComparison.Ordinal) >= 0;
        }

        private static bool EndsWithDelimitedRun(string candidate, string delimiter) =>
            candidate.EndsWith(delimiter, StringComparison.Ordinal)
            && HasEarlierDelimiter(candidate, delimiter, delimiter.Length);

        private static bool HasEarlierDelimiter(string candidate, string delimiter, int closingLength) {
            var closingStart = candidate.Length - closingLength;
            return closingStart > 0
                && candidate.LastIndexOf(delimiter, closingStart - 1, StringComparison.Ordinal) >= 0;
        }

        private static MarkdownReaderOptions CloneOptionsWithoutInlineAutolinks(MarkdownReaderOptions source) {
            var clone = CloneOptionsWithoutFrontMatter(source);
            clone.AutolinkUrls = false;
            clone.AutolinkWwwUrls = false;
            clone.AutolinkBareSchemeUrls = false;
            clone.AutolinkEmails = false;
            return clone;
        }

        private static bool IsFootnoteDefinitionStarter(string line, MarkdownReaderOptions options) {
            if (options?.Footnotes != true || string.IsNullOrWhiteSpace(line)) {
                return false;
            }

            int leading = 0;
            while (leading < line.Length && line[leading] == ' ') {
                leading++;
            }

            if (leading >= 4 || (leading < line.Length && line[leading] == '\t')) {
                return false;
            }

            var trimmed = line.TrimStart();
            if (!(trimmed.Length > 4 && trimmed[0] == '[' && trimmed[1] == '^')) {
                return false;
            }

            int rb = trimmed.IndexOf(']');
            return rb >= 2
                   && rb + 1 < trimmed.Length
                   && trimmed[rb + 1] == ':';
        }

        private static int GetFirstNonWhitespaceColumn(string line) {
            if (string.IsNullOrEmpty(line)) {
                return 1;
            }

            int column = 1;
            for (int i = 0; i < line.Length; i++) {
                char ch = line[i];
                if (ch != ' ' && ch != '\t') {
                    return column;
                }

                column += ch == '\t' ? 4 - ((column - 1) % 4) : 1;
            }

            return column;
        }

        private static int GetTrimmedEndColumn(string line) {
            if (string.IsNullOrEmpty(line)) {
                return 1;
            }

            int endIndex = line.Length - 1;
            while (endIndex >= 0 && char.IsWhiteSpace(line[endIndex])) {
                endIndex--;
            }

            if (endIndex < 0) {
                return 1;
            }

            int column = 1;
            for (int i = 0; i <= endIndex; i++) {
                char ch = line[i];
                if (ch == '\t') {
                    column += 4 - ((column - 1) % 4);
                } else if (i < endIndex) {
                    column++;
                }
            }

            return column;
        }

        private static bool IsReferenceLinkDefinitionStarter(string[] lines, int index, MarkdownReaderOptions options) {
            return TryParseReferenceLinkDefinition(lines, index, options, out _, out _, out _, out _);
        }

        private static bool ShouldTreatReferenceDefinitionAsGenericAttributeParagraph(
            string[] lines,
            int index,
            MarkdownReaderOptions options,
            MarkdownReaderState state) =>
            state?.PendingGenericAttributeBlock != null
            && IsStandaloneAttributeReferenceDefinitionParagraphTarget(lines, index, options);

        private static bool IsAbbreviationDefinitionStarter(string line, MarkdownReaderOptions options) =>
            options?.Abbreviations == true && TryParseAbbreviationDefinition(line, 0, null, out _, out _, out _, out _, out _, out _);

        private static bool IsParagraphTerminatingReferenceLinkDefinition(
            string[] lines,
            int paragraphStartIndex,
            int index,
            MarkdownReaderOptions options,
            MarkdownReaderState state) {
            if (ShouldTreatReferenceDefinitionAsGenericAttributeParagraph(lines, index, options, state)) {
                return false;
            }

            if (!IsReferenceLinkDefinitionStarter(lines, index, options)) {
                return false;
            }

            return index == paragraphStartIndex || CanReferenceDefinitionResolveOpenShortcutParagraph(lines, index);
        }

        private static bool EndsWithTwoSpaces(string s) {
            if (string.IsNullOrEmpty(s)) return false;
            int n = s.Length - 1;
            int count = 0;
            while (n >= 0 && s[n] == ' ') { count++; n--; if (count >= 2) return true; }
            return false;
        }
    }

    private static bool IsQuoteStarter(string line) {
        if (string.IsNullOrEmpty(line)) return false;
        if (CountLeadingIndentColumns(line) > 3) return false;
        var t = line.TrimStart();
        return t.StartsWith(">");
    }

    private static string ExpandReferenceLinks(string text, MarkdownReaderState state) {
        if (state == null || state.LinkRefs.Count == 0 || string.IsNullOrEmpty(text)) return text;
        var sb = new System.Text.StringBuilder(text.Length + 16);
        int pos = 0;
        while (pos < text.Length) {
            if (text[pos] == '[') {
                int rb = FindMatchingBracket(text, pos);
                if (rb > pos + 1) {
                    // collapsed: [text][]
                    if (rb + 2 < text.Length && text[rb + 1] == '[' && text[rb + 2] == ']') {
                        var lbl = text.Substring(pos + 1, rb - (pos + 1));
                        var key = NormalizeReferenceLabel(lbl);
                        if (state.LinkRefs.TryGetValue(key, out var defc)) {
                            sb.Append('[').Append(lbl).Append(']')
                              .Append('(').Append(FormatExpandedReferenceDestination(defc.Url));
                            if (!string.IsNullOrEmpty(defc.Title)) sb.Append(' ').Append('"').Append(defc.Title).Append('"');
                            sb.Append(')');
                            pos = rb + 3; continue;
                        }
                    }
                    // full: [text][label]
                    if (rb + 1 < text.Length && text[rb + 1] == '[') {
                        int rb2 = FindMatchingBracket(text, rb + 1);
                        if (rb2 > rb + 2) {
                            var textLbl = text.Substring(pos + 1, rb - (pos + 1));
                            var refLbl = text.Substring(rb + 2, rb2 - (rb + 2));
                            var key = NormalizeReferenceLabel(refLbl);
                            if (state.LinkRefs.TryGetValue(key, out var def)) {
                                sb.Append('[').Append(textLbl).Append(']')
                                  .Append('(').Append(FormatExpandedReferenceDestination(def.Url));
                                if (!string.IsNullOrEmpty(def.Title)) sb.Append(' ').Append('"').Append(def.Title).Append('"');
                                sb.Append(')');
                                pos = rb2 + 1; continue;
                            }
                        }
                    }
                    // shortcut: [label]
                    if (!(rb + 1 < text.Length && (text[rb + 1] == '(' || text[rb + 1] == '['))) {
                        var lbls = text.Substring(pos + 1, rb - (pos + 1));
                        var key = NormalizeReferenceLabel(lbls);
                        if (state.LinkRefs.TryGetValue(key, out var defs)) {
                            sb.Append('[').Append(lbls).Append(']')
                              .Append('(').Append(FormatExpandedReferenceDestination(defs.Url));
                            if (!string.IsNullOrEmpty(defs.Title)) sb.Append(' ').Append('"').Append(defs.Title).Append('"');
                            sb.Append(')');
                            pos = rb + 1; continue;
                        }
                    }
                }
            }
            sb.Append(text[pos]); pos++;
        }
        return sb.ToString();
    }

    private static string FormatExpandedReferenceDestination(string? url) {
        if (string.IsNullOrEmpty(url)) {
            return string.Empty;
        }

        var value = url!;
        return value.IndexOfAny(new[] { ' ', '\t', '\r', '\n' }) >= 0 ? "<" + value + ">" : value;
    }
}
