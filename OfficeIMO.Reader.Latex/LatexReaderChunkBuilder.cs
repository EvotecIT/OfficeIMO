namespace OfficeIMO.Reader.Latex;

internal static class LatexReaderChunkBuilder {
    internal static IEnumerable<ReaderChunk> BuildBlocks(
        LatexParseResult result,
        string sourceName,
        ReaderOptions readerOptions,
        ReaderLatexOptions options,
        CancellationToken cancellationToken) {
        LatexMarkdownProjection projection = LatexToMarkdownConverter.Project(result.Document, options.MarkdownOptions, cancellationToken);
        var parse = new LatexParseResult(projection.Document, projection.Document.Diagnostics);
        IReadOnlyList<LatexDiagnostic> globalParseDiagnostics = FindUnattachedDiagnostics(parse, projection.Blocks);
        var headingStack = new List<HeadingState>();
        int emitted = 0;
        for (int sourceIndex = 0; sourceIndex < projection.Blocks.Count; sourceIndex++) {
            cancellationToken.ThrowIfCancellationRequested();
            LatexProjectedBlock block = projection.Blocks[sourceIndex];
            if (block.Kind == "heading" && block.Document.Blocks.FirstOrDefault() is HeadingBlock heading) {
                while (headingStack.Count > 0 && headingStack[headingStack.Count - 1].Level >= heading.Level) headingStack.RemoveAt(headingStack.Count - 1);
                headingStack.Add(new HeadingState(heading.Level, heading.Text));
            }
            string markdown = block.Markdown;
            string text = block.Text;
            if (string.IsNullOrWhiteSpace(text)) text = markdown;
            IReadOnlyList<string> parts = Split(text, readerOptions.MaxChars);
            if (parts.Count == 0 && options.IncludeDiagnostics && (block.Diagnostics.Count > 0 ||
                parse.Diagnostics.Any(diagnostic => diagnostic.Span.Start.Offset >= block.Span.Start.Offset && diagnostic.Span.End.Offset <= block.Span.End.Offset))) {
                parts = new[] { string.Empty };
            }
            for (int partIndex = 0; partIndex < parts.Count; partIndex++) {
                bool firstChunk = emitted == 0;
                IReadOnlyList<LatexMarkdownConversionDiagnostic> diagnostics = firstChunk
                    ? projection.GlobalDiagnostics.Concat(block.Diagnostics).ToArray() : block.Diagnostics;
                yield return new ReaderChunk {
                    Id = parts.Count == 1 ? "latex-" + sourceIndex : "latex-" + sourceIndex + "-part-" + (partIndex + 1),
                    Kind = ReaderInputKind.Latex,
                    Location = new ReaderLocation {
                        Path = sourceName, BlockIndex = emitted++, SourceBlockIndex = sourceIndex,
                        StartLine = block.Span.Start.Line, EndLine = InclusiveEnd(block.Span),
                        HeadingPath = headingStack.Count == 0 ? null : string.Join(" > ", headingStack.Select(static item => item.Title)),
                        SourceBlockKind = block.Kind, BlockAnchor = "latex-block-" + sourceIndex
                    },
                    Text = parts[partIndex], Markdown = parts.Count == 1 ? markdown : parts[partIndex],
                    Diagnostics = new ReaderChunkDiagnostics { SourceKind = "latex" },
                    Warnings = options.IncludeDiagnostics ? BuildWarnings(parse, diagnostics, block.Span, parts.Count > 1,
                        firstChunk ? globalParseDiagnostics : null) : null
                };
            }
        }
        string metadataMarkdown = emitted == 0 ? projection.Result.Value.ToMarkdown().TrimEnd() : string.Empty;
        if (emitted == 0 && (!string.IsNullOrWhiteSpace(metadataMarkdown) || options.IncludeDiagnostics && (projection.GlobalDiagnostics.Count > 0 || parse.Diagnostics.Count > 0))) {
            IReadOnlyList<string> parts = Split(metadataMarkdown, readerOptions.MaxChars);
            if (parts.Count == 0) parts = new[] { string.Empty };
            for (int index = 0; index < parts.Count; index++) {
                cancellationToken.ThrowIfCancellationRequested();
                yield return new ReaderChunk {
                    Id = parts.Count == 1 ? "latex-metadata" : "latex-metadata-part-" + (index + 1), Kind = ReaderInputKind.Latex,
                    Location = new ReaderLocation { Path = sourceName, BlockIndex = index, SourceBlockIndex = 0,
                        StartLine = 1, EndLine = projection.Document.Source.LineCount, SourceBlockKind = "metadata", BlockAnchor = "latex-metadata" },
                    Text = parts[index], Markdown = parts[index],
                    Diagnostics = new ReaderChunkDiagnostics { SourceKind = "latex" },
                    Warnings = options.IncludeDiagnostics ? BuildWarnings(parse, projection.GlobalDiagnostics, projection.Document.SyntaxTree.Root.Span, parts.Count > 1) : null
                };
            }
        }
    }

    internal static IEnumerable<ReaderChunk> BuildDocument(
        LatexParseResult result, string sourceName, ReaderOptions readerOptions, ReaderLatexOptions options,
        CancellationToken cancellationToken) {
        LatexMarkdownProjection projection = LatexToMarkdownConverter.Project(result.Document, options.MarkdownOptions, cancellationToken);
        var parse = new LatexParseResult(projection.Document, projection.Document.Diagnostics);
        string markdown = projection.Result.Value.ToMarkdown().TrimEnd();
        string text = string.Join("\n\n", projection.Blocks.Select(static block => block.Text).Where(static text => !string.IsNullOrWhiteSpace(text)));
        if (string.IsNullOrWhiteSpace(text)) text = markdown;
        IReadOnlyList<string> parts = Split(text, readerOptions.MaxChars);
        if (parts.Count == 0 && options.IncludeDiagnostics && (projection.Result.Report.Diagnostics.Count > 0 || parse.Diagnostics.Count > 0)) parts = new[] { string.Empty };
        for (int index = 0; index < parts.Count; index++) {
            cancellationToken.ThrowIfCancellationRequested();
            yield return new ReaderChunk {
                Id = parts.Count == 1 ? "latex-document" : "latex-document-part-" + (index + 1), Kind = ReaderInputKind.Latex,
                Location = new ReaderLocation {
                    Path = sourceName, BlockIndex = index, SourceBlockIndex = 0, StartLine = 1,
                    EndLine = projection.Document.Source.LineCount, SourceBlockKind = "document", BlockAnchor = "latex-document"
                },
                Text = parts[index], Markdown = parts.Count == 1 ? markdown : parts[index],
                Diagnostics = new ReaderChunkDiagnostics { SourceKind = "latex" },
                Warnings = options.IncludeDiagnostics ? BuildWarnings(parse, projection.Result.Report.Diagnostics, projection.Document.SyntaxTree.Root.Span, parts.Count > 1) : null
            };
        }
    }

    private static IReadOnlyList<string>? BuildWarnings(
        LatexParseResult parse,
        IReadOnlyList<LatexMarkdownConversionDiagnostic> conversion,
        LatexSourceSpan span,
        bool split,
        IReadOnlyList<LatexDiagnostic>? globalParseDiagnostics = null) {
        var warnings = parse.Diagnostics.Where(diagnostic => diagnostic.Span.Start.Offset >= span.Start.Offset && diagnostic.Span.End.Offset <= span.End.Offset)
            .Select(diagnostic => diagnostic.Code + ": " + diagnostic.Message).ToList();
        if (globalParseDiagnostics != null) warnings.AddRange(globalParseDiagnostics.Select(diagnostic => diagnostic.Code + ": " + diagnostic.Message));
        if (!parse.Document.IsRecognizedProfile) warnings.Add("LATEXR001: Source is not a recognized OfficeIMO LaTeX article, report, or book profile; preserved structures may be incomplete.");
        warnings.AddRange(conversion.Select(diagnostic => diagnostic.Code + ": " + diagnostic.Message));
        if (split) warnings.Add("LaTeX content was split due to ReaderOptions.MaxChars.");
        return warnings.Count == 0 ? null : warnings;
    }

    private static IReadOnlyList<LatexDiagnostic> FindUnattachedDiagnostics(LatexParseResult parse, IReadOnlyList<LatexProjectedBlock> blocks) {
        var unattached = new List<LatexDiagnostic>();
        int index = 0;
        foreach (LatexDiagnostic diagnostic in parse.Diagnostics.OrderBy(static item => item.Span.Start.Offset)) {
            while (index < blocks.Count && blocks[index].Span.End.Offset < diagnostic.Span.Start.Offset) index++;
            if (index == blocks.Count || diagnostic.Span.Start.Offset < blocks[index].Span.Start.Offset || diagnostic.Span.End.Offset > blocks[index].Span.End.Offset) {
                unattached.Add(diagnostic);
            }
        }
        return unattached;
    }

    private static IReadOnlyList<string> Split(string value, int maximum) {
        if (value.Length == 0) return Array.Empty<string>();
        if (maximum <= 0 || value.Length <= maximum) return new[] { value };
        var parts = new List<string>();
        int offset = 0;
        while (offset < value.Length) {
            int length = Math.Min(maximum, value.Length - offset);
            int end = offset + length;
            if (end < value.Length) {
                int split = value.LastIndexOf('\n', end - 1, length);
                if (split <= offset) split = value.LastIndexOf(' ', end - 1, length);
                if (split > offset) length = split - offset;
            }
            parts.Add(value.Substring(offset, length).Trim());
            offset += length;
            while (offset < value.Length && char.IsWhiteSpace(value[offset])) offset++;
        }
        return parts;
    }

    private static int InclusiveEnd(LatexSourceSpan span) => span.End.Column == 1 && span.End.Line > span.Start.Line ? span.End.Line - 1 : span.End.Line;

    private sealed class HeadingState {
        internal HeadingState(int level, string title) { Level = level; Title = title; }
        internal int Level { get; }
        internal string Title { get; }
    }
}
