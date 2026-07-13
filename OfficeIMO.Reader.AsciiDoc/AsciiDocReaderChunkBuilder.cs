namespace OfficeIMO.Reader.AsciiDoc;

internal static class AsciiDocReaderChunkBuilder {
    internal static IEnumerable<ReaderChunk> BuildBlockChunks(
        AsciiDocParseResult result,
        string sourceName,
        ReaderOptions readerOptions,
        ReaderAsciiDocOptions options,
        CancellationToken cancellationToken) {
        var headingStack = new List<HeadingState>();
        var attachedBlocks = new HashSet<AsciiDocBlock>(
            result.Document.BlocksOfType<AsciiDocListBlock>()
                .SelectMany(static list => list.Items)
                .SelectMany(static item => item.AttachedBlocks));
        AsciiDocDocumentAttributes attributes = result.Document.GetAttributes();
        int emittedIndex = 0;
        for (int sourceIndex = 0; sourceIndex < result.Document.Blocks.Count; sourceIndex++) {
            cancellationToken.ThrowIfCancellationRequested();
            AsciiDocBlock block = result.Document.Blocks[sourceIndex];
            if (attachedBlocks.Contains(block)) continue;
            if (!ShouldEmit(block, options)) continue;

            if (block is AsciiDocHeading heading) UpdateHeadingStack(headingStack, heading);
            string headingPath = string.Join(" > ", headingStack.Select(static state => state.Title));
            string text = GetPlainText(block);
            AsciiDocToMarkdownResult markdownResult = AsciiDocToMarkdownConverter.ConvertBlock(block, attributes, options.MarkdownOptions);
            string markdown = markdownResult.Value.ToMarkdown().TrimEnd();
            if (markdown.Length == 0 && block is AsciiDocAttributeEntry) markdown = block.OriginalText.TrimEnd('\r', '\n');

            IReadOnlyList<string> parts = Split(text.Length == 0 ? markdown : text, readerOptions.MaxChars);
            if (parts.Count == 0) parts = new[] { string.Empty };
            for (int partIndex = 0; partIndex < parts.Count; partIndex++) {
                IReadOnlyList<string>? warnings = options.IncludeDiagnostics
                    ? BuildWarnings(result.Diagnostics, markdownResult.Report.Diagnostics, block, parts.Count > 1)
                    : null;
                yield return new ReaderChunk {
                    Id = BuildId(sourceIndex, partIndex, parts.Count),
                    Kind = ReaderInputKind.AsciiDoc,
                    Location = new ReaderLocation {
                        Path = sourceName,
                        BlockIndex = emittedIndex++,
                        SourceBlockIndex = sourceIndex,
                        StartLine = block.Span.Start.Line,
                        EndLine = GetInclusiveEndLine(block.Span),
                        HeadingPath = headingPath.Length == 0 ? null : headingPath,
                        SourceBlockKind = GetBlockKind(block),
                        BlockAnchor = "asciidoc-block-" + sourceIndex.ToString(System.Globalization.CultureInfo.InvariantCulture)
                    },
                    Text = parts[partIndex],
                    Markdown = parts.Count == 1 ? markdown : parts[partIndex],
                    Diagnostics = new ReaderChunkDiagnostics { SourceKind = "asciidoc" },
                    Warnings = warnings
                };
            }
        }
    }

    internal static IEnumerable<ReaderChunk> BuildDocumentChunks(
        AsciiDocParseResult result,
        string sourceName,
        ReaderOptions readerOptions,
        ReaderAsciiDocOptions options,
        CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        AsciiDocToMarkdownResult conversion = result.Document.ToMarkdownDocumentResult(options.MarkdownOptions);
        var attachedBlocks = new HashSet<AsciiDocBlock>(
            result.Document.BlocksOfType<AsciiDocListBlock>()
                .SelectMany(static list => list.Items)
                .SelectMany(static item => item.AttachedBlocks));
        string text = string.Join("\n\n", result.Document.Blocks
            .Where(block => !attachedBlocks.Contains(block) && ShouldEmit(block, options))
            .Select(GetPlainText)
            .Where(value => value.Length > 0));
        string markdown = conversion.Value.ToMarkdown().TrimEnd();
        IReadOnlyList<string> parts = Split(text.Length == 0 ? markdown : text, readerOptions.MaxChars);
        if (parts.Count == 0) parts = new[] { string.Empty };

        for (int partIndex = 0; partIndex < parts.Count; partIndex++) {
            cancellationToken.ThrowIfCancellationRequested();
            yield return new ReaderChunk {
                Id = BuildId(0, partIndex, parts.Count),
                Kind = ReaderInputKind.AsciiDoc,
                Location = new ReaderLocation {
                    Path = sourceName,
                    BlockIndex = partIndex,
                    SourceBlockIndex = 0,
                    StartLine = 1,
                    EndLine = GetDocumentEndLine(result.Document.Source),
                    SourceBlockKind = "document",
                    BlockAnchor = "asciidoc-document"
                },
                Text = parts[partIndex],
                Markdown = parts.Count == 1 ? markdown : parts[partIndex],
                Diagnostics = new ReaderChunkDiagnostics { SourceKind = "asciidoc" },
                Warnings = options.IncludeDiagnostics
                    ? BuildWarnings(result.Diagnostics, conversion.Report.Diagnostics, null, parts.Count > 1)
                    : null
            };
        }
    }

    private static bool ShouldEmit(AsciiDocBlock block, ReaderAsciiDocOptions options) {
        if (block is AsciiDocBlankLine) return false;
        if (block is AsciiDocLineComment) return options.IncludeComments;
        if (block is AsciiDocDelimitedBlock delimited && delimited.Kind == AsciiDocDelimitedBlockKind.Comment) return options.IncludeComments;
        if (block is AsciiDocAttributeEntry) return options.IncludeAttributes;
        if (block is IAsciiDocBlockMetadata || block is AsciiDocListContinuation) return false;
        return true;
    }

    private static string GetPlainText(AsciiDocBlock block) {
        switch (block) {
            case AsciiDocHeading heading: return heading.Title;
            case AsciiDocParagraph paragraph: return paragraph.Text;
            case AsciiDocListBlock list:
                return string.Join("\n", list.Items.Select(item => string.Join("\n",
                    new[] { item.Text }.Concat(item.AttachedBlocks.Select(GetPlainText)).Where(static value => value.Length > 0))));
            case AsciiDocDescriptionListBlock list: return string.Join("\n", list.Items.Select(static item => item.Term + ": " + item.Description));
            case AsciiDocAdmonitionBlock admonition: return admonition.Label + ": " + admonition.Text;
            case AsciiDocTableBlock table: return string.Join("\n", table.Table.Rows.Select(row => string.Join("\t", row.Cells.Select(static cell => cell.Value))));
            case AsciiDocDelimitedBlock delimited: return delimited.Content.TrimEnd('\r', '\n');
            case AsciiDocLineComment comment: return comment.Text;
            case AsciiDocAttributeEntry attribute: return attribute.Name + (attribute.Value.Length == 0 ? string.Empty : ": " + attribute.Value);
            default: return block.OriginalText.TrimEnd('\r', '\n');
        }
    }

    private static string GetBlockKind(AsciiDocBlock block) {
        if (block is AsciiDocHeading) return "heading";
        if (block is AsciiDocParagraph) return "paragraph";
        if (block is AsciiDocListBlock list) return list.Kind == AsciiDocListKind.Ordered ? "ordered-list" : "unordered-list";
        if (block is AsciiDocDescriptionListBlock) return "description-list";
        if (block is AsciiDocAdmonitionBlock) return "admonition";
        if (block is AsciiDocTableBlock) return "table";
        if (block is AsciiDocDelimitedBlock delimited) return "delimited-" + delimited.Kind.ToString().ToLowerInvariant();
        if (block is AsciiDocBlockMacro) return "block-macro";
        if (block is AsciiDocAttributeEntry) return "attribute";
        if (block is AsciiDocLineComment) return "comment";
        return "raw";
    }

    private static void UpdateHeadingStack(List<HeadingState> stack, AsciiDocHeading heading) {
        int level = heading.IsDocumentTitle ? 0 : heading.SectionLevel;
        while (stack.Count > 0 && stack[stack.Count - 1].Level >= level) stack.RemoveAt(stack.Count - 1);
        stack.Add(new HeadingState(level, heading.Title));
    }

    private static IReadOnlyList<string>? BuildWarnings(
        IReadOnlyList<AsciiDocDiagnostic> parserDiagnostics,
        IReadOnlyList<AsciiDocMarkdownConversionDiagnostic> conversionDiagnostics,
        AsciiDocBlock? block,
        bool wasSplit) {
        var warnings = new List<string>();
        for (int index = 0; index < parserDiagnostics.Count; index++) {
            AsciiDocDiagnostic diagnostic = parserDiagnostics[index];
            if (block == null || block.Span.Contains(diagnostic.Span)) warnings.Add(diagnostic.Code + ": " + diagnostic.Message);
        }
        for (int index = 0; index < conversionDiagnostics.Count; index++) {
            AsciiDocMarkdownConversionDiagnostic diagnostic = conversionDiagnostics[index];
            warnings.Add(diagnostic.Code + ": " + diagnostic.Message);
        }
        if (wasSplit) warnings.Add("AsciiDoc content was split due to ReaderOptions.MaxChars.");
        return warnings.Count == 0 ? null : warnings;
    }

    private static IReadOnlyList<string> Split(string value, int maxChars) {
        if (value.Length == 0) return Array.Empty<string>();
        if (maxChars <= 0 || value.Length <= maxChars) return new[] { value };
        var parts = new List<string>();
        int offset = 0;
        while (offset < value.Length) {
            int length = Math.Min(maxChars, value.Length - offset);
            int end = offset + length;
            if (end < value.Length) {
                int breakAt = value.LastIndexOf('\n', end - 1, length);
                if (breakAt <= offset) breakAt = value.LastIndexOf(' ', end - 1, length);
                if (breakAt > offset) length = breakAt - offset;
            }
            parts.Add(value.Substring(offset, length).Trim());
            offset += length;
            while (offset < value.Length && char.IsWhiteSpace(value[offset])) offset++;
        }
        return parts;
    }

    private static string BuildId(int blockIndex, int partIndex, int partCount) =>
        partCount <= 1
            ? "asciidoc-" + blockIndex.ToString(System.Globalization.CultureInfo.InvariantCulture)
            : "asciidoc-" + blockIndex.ToString(System.Globalization.CultureInfo.InvariantCulture) + "-part-" + (partIndex + 1).ToString(System.Globalization.CultureInfo.InvariantCulture);

    private static int GetInclusiveEndLine(AsciiDocSourceSpan span) =>
        span.End.Column == 1 && span.End.Line > span.Start.Line ? span.End.Line - 1 : span.End.Line;

    private static int GetDocumentEndLine(AsciiDocSourceText source) =>
        source.Text.Length == 0 ? 1 : source.GetPosition(source.Text.Length - 1).Line;

    private sealed class HeadingState {
        internal HeadingState(int level, string title) { Level = level; Title = title; }
        internal int Level { get; }
        internal string Title { get; }
    }
}
