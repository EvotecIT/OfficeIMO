namespace OfficeIMO.Markdown;

public static partial class MarkdownReader {
    private static (IReadOnlyList<IMarkdownBlock> Blocks, IReadOnlyList<MarkdownSyntaxNode> SyntaxChildren) PreserveMarkdigDefinitionLazyParagraphSoftBreaks(
        IReadOnlyList<IMarkdownBlock> blocks,
        IReadOnlyList<MarkdownSyntaxNode> syntaxChildren,
        IReadOnlyList<MarkdownSourceLineSlice> sourceLines,
        MarkdownReaderOptions options,
        MarkdownReaderState state) {
        if (blocks == null || syntaxChildren == null || sourceLines == null ||
            blocks.Count == 0 || blocks.Count != syntaxChildren.Count) {
            return (blocks ?? Array.Empty<IMarkdownBlock>(), syntaxChildren ?? Array.Empty<MarkdownSyntaxNode>());
        }

        List<MarkdownSyntaxNode>? updatedSyntax = null;
        for (int index = 0; index < blocks.Count; index++) {
            if (blocks[index] is not ParagraphBlock paragraph ||
                syntaxChildren[index].Kind != MarkdownSyntaxKind.Paragraph ||
                !TryCreateDefinitionLazyParagraphInlines(
                    syntaxChildren[index],
                    sourceLines,
                    options,
                    state,
                    out var preservedInlines)) {
                continue;
            }

            paragraph.Inlines.AutoSpacing = false;
            paragraph.Inlines.ReplaceItems(preservedInlines.Nodes);
            updatedSyntax ??= new List<MarkdownSyntaxNode>(syntaxChildren);
            updatedSyntax[index] = MarkdownBlockSyntaxBuilder.BuildInlineBlock(paragraph, syntaxChildren[index].SourceSpan);
        }

        return updatedSyntax == null
            ? (blocks, syntaxChildren)
            : (blocks, updatedSyntax);
    }

    private static bool TryCreateDefinitionLazyParagraphInlines(
        MarkdownSyntaxNode paragraphSyntax,
        IReadOnlyList<MarkdownSourceLineSlice> sourceLines,
        MarkdownReaderOptions options,
        MarkdownReaderState state,
        out InlineSequence inlines) {
        inlines = new InlineSequence { AutoSpacing = false };
        if (paragraphSyntax == null ||
            !paragraphSyntax.SourceSpan.HasValue ||
            sourceLines == null ||
            sourceLines.Count < 2) {
            return false;
        }

        var paragraphSourceLines = GetDefinitionParagraphSourceLines(paragraphSyntax.SourceSpan.Value, sourceLines);
        if (paragraphSourceLines.Count < 2 || !HasDefinitionLazyParagraphContinuationLine(paragraphSourceLines)) {
            return false;
        }

        inlines = ParseDefinitionLazyParagraphInlines(paragraphSourceLines, options, state);
        return true;
    }

    private static List<MarkdownSourceLineSlice> GetDefinitionParagraphSourceLines(
        MarkdownSourceSpan span,
        IReadOnlyList<MarkdownSourceLineSlice> sourceLines) {
        var lines = new List<MarkdownSourceLineSlice>();
        for (int i = 0; i < sourceLines.Count; i++) {
            var sourceLine = sourceLines[i];
            if (sourceLine.AbsoluteLine < span.StartLine || sourceLine.AbsoluteLine > span.EndLine) {
                continue;
            }

            if (string.IsNullOrEmpty(sourceLine.Text)) {
                continue;
            }

            lines.Add(sourceLine);
        }

        return lines;
    }

    private static bool HasDefinitionLazyParagraphContinuationLine(IReadOnlyList<MarkdownSourceLineSlice> lines) {
        if (lines == null || lines.Count < 2) {
            return false;
        }

        MarkdownSourceLineSlice? firstContentLine = null;
        for (int i = 0; i < lines.Count; i++) {
            if (string.IsNullOrWhiteSpace(lines[i].Text)) {
                continue;
            }

            firstContentLine = lines[i];
            break;
        }

        if (!firstContentLine.HasValue) {
            return false;
        }

        for (int i = 0; i < lines.Count; i++) {
            var line = lines[i];
            if (line.AbsoluteLine <= firstContentLine.Value.AbsoluteLine ||
                string.IsNullOrWhiteSpace(line.Text)) {
                continue;
            }

            if (line.StartColumn < firstContentLine.Value.StartColumn) {
                return true;
            }
        }

        return false;
    }

    private static InlineSequence ParseDefinitionLazyParagraphInlines(
        IReadOnlyList<MarkdownSourceLineSlice> lines,
        MarkdownReaderOptions options,
        MarkdownReaderState state) {
        if (state?.SourceTextMap == null) {
            return ParseDefinitionLazyParagraphInlinesWithoutSourceMap(lines, options, state);
        }

        var (text, sourceMap) = JoinDefinitionLazyParagraphLinesWithSourceMap(lines, options, state);
        return ParseInlines(text, options, state, sourceMap);
    }

    private static InlineSequence ParseDefinitionLazyParagraphInlinesWithoutSourceMap(
        IReadOnlyList<MarkdownSourceLineSlice> lines,
        MarkdownReaderOptions options,
        MarkdownReaderState? state) {
        var sequence = new InlineSequence { AutoSpacing = false };
        for (int i = 0; i < lines.Count; i++) {
            if (i > 0) {
                sequence.AddRaw(new SoftBreakInline());
            }

            var parsedLine = ParseInlines(lines[i].Text, options, state);
            for (int nodeIndex = 0; nodeIndex < parsedLine.Nodes.Count; nodeIndex++) {
                sequence.AddRaw(parsedLine.Nodes[nodeIndex]);
            }
        }

        return sequence;
    }

    private static (string Text, MarkdownInlineSourceMap SourceMap) JoinDefinitionLazyParagraphLinesWithSourceMap(
        IReadOnlyList<MarkdownSourceLineSlice> lines,
        MarkdownReaderOptions options,
        MarkdownReaderState state) {
        var text = new StringBuilder();
        var points = new List<MarkdownSourcePoint?>();
        var tokenSpans = new List<MarkdownSourceSpan?>();
        var tokenLiterals = new List<string?>();
        MarkdownSourceLineSlice? previousLine = null;
        ParagraphLineJoinInfo? previousJoinInfo = null;

        for (int i = 0; i < lines.Count; i++) {
            var line = lines[i];
            if (previousLine.HasValue && previousJoinInfo != null) {
                AppendDefinitionLazyParagraphLineBreak(
                    text,
                    points,
                    tokenSpans,
                    tokenLiterals,
                    previousLine.Value,
                    previousJoinInfo.Value,
                    state);
            }

            var joinInfo = GetParagraphLineJoinInfo(
                line.Text,
                line.AbsoluteLine,
                line.StartColumn,
                options,
                state.SourceTextMap,
                hasFollowingLine: i + 1 < lines.Count);
            text.Append(joinInfo.Text);
            for (int charIndex = 0; charIndex < joinInfo.Text.Length; charIndex++) {
                points.Add(state.SourceTextMap!.CreatePoint(line.AbsoluteLine, line.StartColumn + charIndex));
                tokenSpans.Add(null);
                tokenLiterals.Add(null);
            }

            previousLine = line;
            previousJoinInfo = joinInfo;
        }

        return (text.ToString(), new MarkdownInlineSourceMap(points.ToArray(), tokenSpans.ToArray(), tokenLiterals.ToArray()));
    }

    private static void AppendDefinitionLazyParagraphLineBreak(
        StringBuilder text,
        List<MarkdownSourcePoint?> points,
        List<MarkdownSourceSpan?> tokenSpans,
        List<string?> tokenLiterals,
        MarkdownSourceLineSlice previousLine,
        ParagraphLineJoinInfo previousJoinInfo,
        MarkdownReaderState state) {
        text.Append('\n');
        var previousEndColumn = previousLine.StartColumn + Math.Max(0, previousJoinInfo.Text.Length - 1);
        points.Add(state.SourceTextMap?.CreatePoint(previousLine.AbsoluteLine, previousEndColumn));
        if (previousJoinInfo.HardBreak) {
            tokenSpans.Add(previousJoinInfo.HardBreakMarkerSpan);
            tokenLiterals.Add(previousJoinInfo.HardBreakMarker);
            return;
        }

        tokenSpans.Add(CreateSpan(
            state,
            previousLine.AbsoluteLine,
            previousEndColumn,
            previousLine.AbsoluteLine,
            previousEndColumn));
        tokenLiterals.Add("\n");
    }
}
