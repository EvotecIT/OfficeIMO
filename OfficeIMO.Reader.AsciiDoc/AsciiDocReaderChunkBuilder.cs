namespace OfficeIMO.Reader.AsciiDoc;

internal static class AsciiDocReaderChunkBuilder {
    internal static IEnumerable<ReaderChunk> BuildBlockChunks(
        AsciiDocParseResult result,
        string sourceName,
        ReaderOptions readerOptions,
        ReaderAsciiDocOptions options,
        CancellationToken cancellationToken) {
        // Options belong to this read operation. Reuse one document catalog so
        // cross-block references resolve without rebuilding it for every chunk.
        options.MarkdownOptions.References ??= AsciiDocReferenceCatalog.Create(result.Document, options.MarkdownOptions.MaximumBlockNestingDepth, cancellationToken);
        var headingStack = new List<HeadingState>();
        var attachedBlocks = new HashSet<AsciiDocBlock>(
            result.Document.BlocksOfType<AsciiDocListBlock>()
                .SelectMany(static list => list.Items)
                .SelectMany(static item => item.AttachedBlocks));
        int emittedIndex = 0;
        int sourceIndex = -1;
        foreach (AsciiDocBlockContext context in result.Document.GetBlockContexts(null, true, options.MarkdownOptions.MaximumBlockNestingDepth, cancellationToken)) {
            sourceIndex++;
            cancellationToken.ThrowIfCancellationRequested();
            AsciiDocBlock block = context.Block;
            if (attachedBlocks.Contains(block)) continue;
            if (!ShouldEmit(block, options)) continue;

            if (block is AsciiDocHeading heading) UpdateHeadingStack(headingStack, heading, ResolveText(heading.Title, context.Attributes, options));
            string headingPath = string.Join(" > ", headingStack.Select(static state => state.Title));
            string text = GetPlainText(block, context.Attributes, options, cancellationToken);
            AsciiDocToMarkdownResult markdownResult = block.ToMarkdownDocumentResult(context.Attributes, options.MarkdownOptions);
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
        string text = string.Join("\n\n", result.Document.GetBlockContexts(null, true, options.MarkdownOptions.MaximumBlockNestingDepth, cancellationToken)
            .Where(context => !attachedBlocks.Contains(context.Block) && ShouldEmit(context.Block, options))
            .Select(context => GetPlainText(context.Block, context.Attributes, options, cancellationToken))
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

    private static string GetPlainText(AsciiDocBlock block, AsciiDocDocumentAttributes attributes, ReaderAsciiDocOptions options, CancellationToken token, int depth = 0) {
        token.ThrowIfCancellationRequested();
        if (depth >= options.MarkdownOptions.MaximumBlockNestingDepth) throw new InvalidDataException("AsciiDoc Reader text exceeds MaximumBlockNestingDepth.");
        string Resolve(string text) => ResolveText(text, attributes, options);
        switch (block) {
            case AsciiDocHeading heading: return Resolve(heading.Title);
            case AsciiDocParagraph paragraph: return Resolve(paragraph.Text);
            case AsciiDocListBlock list:
                return string.Join("\n", list.Items.Select(item => string.Join("\n",
                    new[] { Resolve(item.Text) }.Concat(item.AttachedBlocks.Select(child => GetPlainText(child, attributes, options, token, depth + 1))).Where(static value => value.Length > 0))));
            case AsciiDocDescriptionListBlock list: return string.Join("\n", list.Items.Select(item => Resolve(item.Term) + ": " + Resolve(item.Description)));
            case AsciiDocAdmonitionBlock admonition: return admonition.Label + ": " + Resolve(admonition.Text);
            case AsciiDocTableBlock table: return string.Join("\n", table.Table.Rows.Select(row => string.Join("\t", row.Cells.Select(cell => Resolve(cell.Value)))));
            case AsciiDocDelimitedBlock compound when compound.GetBody(token) is AsciiDocDocument body:
                int remaining = options.MarkdownOptions.MaximumBlockNestingDepth - depth - 1;
                if (remaining < 1) throw new InvalidDataException("AsciiDoc Reader text exceeds MaximumBlockNestingDepth.");
                var attached = new HashSet<AsciiDocBlock>(body.BlocksOfType<AsciiDocListBlock>().SelectMany(list => list.Items).SelectMany(item => item.AttachedBlocks));
                return string.Join("\n\n", body.GetBlockContextsFromSnapshot(attributes, true, remaining, token)
                    .Where(context => !attached.Contains(context.Block) && ShouldEmit(context.Block, options)).Select(context => GetPlainText(context.Block, context.Attributes, options, token, depth + 1)).Where(value => value.Length > 0));
            case AsciiDocDelimitedBlock delimited: return delimited.Content.TrimEnd('\r', '\n');
            case AsciiDocLineComment comment: return comment.Text;
            case AsciiDocAttributeEntry attribute: return attribute.Name + (attribute.Value.Length == 0 ? string.Empty : ": " + attribute.Value);
            default: return block.OriginalText.TrimEnd('\r', '\n');
        }
    }

    private static string ResolveText(string text, AsciiDocDocumentAttributes attributes, ReaderAsciiDocOptions options) =>
        options.MarkdownOptions.ExpandDocumentAttributes ? AsciiDocAttributeSubstitutor.Substitute(text, attributes,
            new AsciiDocAttributeSubstitutionOptions { UndefinedAttributeBehavior = options.MarkdownOptions.UndefinedAttributeBehavior }).Value : text;

    private static string GetBlockKind(AsciiDocBlock block) {
        if (block is AsciiDocHeading) return "heading";
        if (block is AsciiDocParagraph) return "paragraph";
        if (block is AsciiDocListBlock list) return list.Kind == AsciiDocListKind.Callout ? "callout-list" : list.Kind == AsciiDocListKind.Ordered ? "ordered-list" : "unordered-list";
        if (block is AsciiDocDescriptionListBlock) return "description-list";
        if (block is AsciiDocAdmonitionBlock) return "admonition";
        if (block is AsciiDocTableBlock) return "table";
        if (block is AsciiDocDelimitedBlock delimited) return "delimited-" + delimited.Kind.ToString().ToLowerInvariant();
        if (block is AsciiDocBlockMacro) return "block-macro";
        if (block is AsciiDocAttributeEntry) return "attribute";
        if (block is AsciiDocLineComment) return "comment";
        return "raw";
    }

    private static void UpdateHeadingStack(List<HeadingState> stack, AsciiDocHeading heading, string title) {
        int level = heading.IsDocumentTitle ? 0 : heading.SectionLevel;
        while (stack.Count > 0 && stack[stack.Count - 1].Level >= level) stack.RemoveAt(stack.Count - 1);
        stack.Add(new HeadingState(level, title));
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
