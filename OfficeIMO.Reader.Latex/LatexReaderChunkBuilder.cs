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
        var parseByBlock = new List<LatexDiagnostic>[projection.Blocks.Count];
        for (int index = 0; index < parseByBlock.Length; index++) parseByBlock[index] = new List<LatexDiagnostic>();
        IReadOnlyList<LatexDiagnostic> globalParseDiagnostics = PartitionDiagnostics(parse, projection.Blocks, parseByBlock, cancellationToken);
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
            IReadOnlyList<LatexTextSegment> content = block.TextSegments(cancellationToken);
            if (content.All(static segment => string.IsNullOrWhiteSpace(segment.Text)) && !string.IsNullOrWhiteSpace(markdown))
                content = new[] { new LatexTextSegment(markdown, "markdown") };
            IReadOnlyList<LatexReaderPart> parts = LatexReaderTextSplitter.Split(content, readerOptions.MaxChars, cancellationToken);
            if (parts.Count == 0 && options.IncludeDiagnostics && (block.Diagnostics.Count > 0 ||
                parseByBlock[sourceIndex].Count > 0)) {
                parts = new[] { new LatexReaderPart(string.Empty, string.Empty) };
            }
            bool firstChunk = emitted == 0;
            IReadOnlyList<LatexMarkdownConversionDiagnostic> diagnostics = firstChunk
                ? projection.GlobalDiagnostics.Concat(block.Diagnostics).ToArray() : block.Diagnostics;
            IReadOnlyList<string>? warnings = options.IncludeDiagnostics ? BuildWarnings(parse.Document.IsRecognizedProfile, parseByBlock[sourceIndex], block.Diagnostics, parts.Count > 1, cancellationToken) : null;
            IReadOnlyList<string>? firstWarnings = options.IncludeDiagnostics && firstChunk
                ? BuildWarnings(parse.Document.IsRecognizedProfile, parseByBlock[sourceIndex], diagnostics, parts.Count > 1, cancellationToken, globalParseDiagnostics) : warnings;
            for (int partIndex = 0; partIndex < parts.Count; partIndex++) {
                cancellationToken.ThrowIfCancellationRequested();
                yield return new ReaderChunk {
                    Id = parts.Count == 1 ? "latex-" + sourceIndex : "latex-" + sourceIndex + "-part-" + (partIndex + 1),
                    Kind = ReaderInputKind.Latex,
                    Location = new ReaderLocation {
                        Path = sourceName, BlockIndex = emitted++, SourceBlockIndex = sourceIndex,
                        StartLine = block.Span.Start.Line, EndLine = InclusiveEnd(block.Span),
                        HeadingPath = headingStack.Count == 0 ? null : string.Join(" > ", headingStack.Select(static item => item.Title)),
                        SourceBlockKind = block.Kind, BlockAnchor = "latex-block-" + sourceIndex
                    },
                    Text = parts[partIndex].Text, Markdown = parts.Count == 1 ? markdown : parts[partIndex].Markdown,
                    Diagnostics = new ReaderChunkDiagnostics { SourceKind = "latex" },
                    Warnings = partIndex == 0 ? firstWarnings : warnings
                };
            }
        }
        string metadataMarkdown = emitted == 0 ? projection.Result.Value.ToMarkdown().TrimEnd() : string.Empty;
        if (emitted == 0 && (!string.IsNullOrWhiteSpace(metadataMarkdown) || options.IncludeDiagnostics && (projection.GlobalDiagnostics.Count > 0 || parse.Diagnostics.Count > 0))) {
            IReadOnlyList<LatexReaderPart> parts = LatexReaderTextSplitter.Split(new[] { new LatexTextSegment(metadataMarkdown) }, readerOptions.MaxChars, cancellationToken);
            IReadOnlyList<string>? firstWarnings = options.IncludeDiagnostics ? BuildWarnings(parse.Document.IsRecognizedProfile, parse.Diagnostics, projection.GlobalDiagnostics, parts.Count > 1, cancellationToken) : null;
            IReadOnlyList<string>? splitWarnings = options.IncludeDiagnostics ? BuildWarnings(true, Array.Empty<LatexDiagnostic>(), Array.Empty<LatexMarkdownConversionDiagnostic>(), parts.Count > 1, cancellationToken) : null;
            if (parts.Count == 0) parts = new[] { new LatexReaderPart(string.Empty, string.Empty) };
            for (int index = 0; index < parts.Count; index++) {
                cancellationToken.ThrowIfCancellationRequested();
                yield return new ReaderChunk {
                    Id = parts.Count == 1 ? "latex-metadata" : "latex-metadata-part-" + (index + 1), Kind = ReaderInputKind.Latex,
                    Location = new ReaderLocation { Path = sourceName, BlockIndex = index, SourceBlockIndex = 0,
                        StartLine = 1, EndLine = projection.Document.Source.LineCount, SourceBlockKind = "metadata", BlockAnchor = "latex-metadata" },
                    Text = parts[index].Text, Markdown = parts.Count == 1 ? metadataMarkdown : parts[index].Markdown,
                    Diagnostics = new ReaderChunkDiagnostics { SourceKind = "latex" },
                    Warnings = index == 0 ? firstWarnings : splitWarnings
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
        var segments = new List<LatexTextSegment>();
        foreach (LatexProjectedBlock block in projection.Blocks) {
            cancellationToken.ThrowIfCancellationRequested();
            IReadOnlyList<LatexTextSegment> content = block.TextSegments(cancellationToken);
            if (content.All(static segment => string.IsNullOrWhiteSpace(segment.Text))) continue;
            if (segments.Count > 0) segments.Add(new LatexTextSegment("\n\n"));
            segments.AddRange(content);
        }
        if (segments.Count == 0) segments.Add(new LatexTextSegment(markdown));
        IReadOnlyList<LatexReaderPart> parts = LatexReaderTextSplitter.Split(segments, readerOptions.MaxChars, cancellationToken);
        IReadOnlyList<string>? firstWarnings = options.IncludeDiagnostics ? BuildWarnings(parse.Document.IsRecognizedProfile, parse.Diagnostics, projection.Result.Report.Diagnostics, parts.Count > 1, cancellationToken) : null;
        IReadOnlyList<string>? splitWarnings = options.IncludeDiagnostics ? BuildWarnings(true, Array.Empty<LatexDiagnostic>(), Array.Empty<LatexMarkdownConversionDiagnostic>(), parts.Count > 1, cancellationToken) : null;
        if (parts.Count == 0 && options.IncludeDiagnostics && (projection.Result.Report.Diagnostics.Count > 0 || parse.Diagnostics.Count > 0)) parts = new[] { new LatexReaderPart(string.Empty, string.Empty) };
        for (int index = 0; index < parts.Count; index++) {
            cancellationToken.ThrowIfCancellationRequested();
            yield return new ReaderChunk {
                Id = parts.Count == 1 ? "latex-document" : "latex-document-part-" + (index + 1), Kind = ReaderInputKind.Latex,
                Location = new ReaderLocation {
                    Path = sourceName, BlockIndex = index, SourceBlockIndex = 0, StartLine = 1,
                    EndLine = projection.Document.Source.LineCount, SourceBlockKind = "document", BlockAnchor = "latex-document"
                },
                Text = parts[index].Text, Markdown = parts.Count == 1 ? markdown : parts[index].Markdown,
                Diagnostics = new ReaderChunkDiagnostics { SourceKind = "latex" },
                Warnings = index == 0 ? firstWarnings : splitWarnings
            };
        }
    }

    private static IReadOnlyList<string>? BuildWarnings(
        bool recognizedProfile,
        IReadOnlyList<LatexDiagnostic> localParseDiagnostics,
        IReadOnlyList<LatexMarkdownConversionDiagnostic> conversion,
        bool split,
        CancellationToken cancellationToken,
        IReadOnlyList<LatexDiagnostic>? globalParseDiagnostics = null) {
        var warnings = new List<string>();
        foreach (LatexDiagnostic diagnostic in localParseDiagnostics) {
            cancellationToken.ThrowIfCancellationRequested();
            warnings.Add(diagnostic.Code + ": " + diagnostic.Message);
        }
        if (globalParseDiagnostics != null) {
            foreach (LatexDiagnostic diagnostic in globalParseDiagnostics) {
                cancellationToken.ThrowIfCancellationRequested();
                warnings.Add(diagnostic.Code + ": " + diagnostic.Message);
            }
        }
        if (!recognizedProfile) warnings.Add("LATEXR001: Source is not a recognized OfficeIMO LaTeX article, report, or book profile; preserved structures may be incomplete.");
        foreach (LatexMarkdownConversionDiagnostic diagnostic in conversion) {
            cancellationToken.ThrowIfCancellationRequested();
            warnings.Add(diagnostic.Code + ": " + diagnostic.Message);
        }
        if (split) warnings.Add("LaTeX content was split due to ReaderOptions.MaxChars; split Markdown flattens layout and formatting while retaining literal text, code, and label anchors.");
        return warnings.Count == 0 ? null : warnings;
    }

    // Assign each parser diagnostic once. Repeated chunks reuse their block warnings.
    private static IReadOnlyList<LatexDiagnostic> PartitionDiagnostics(LatexParseResult parse, IReadOnlyList<LatexProjectedBlock> blocks,
        List<LatexDiagnostic>[] byBlock, CancellationToken cancellationToken) {
        var unattached = new List<LatexDiagnostic>();
        foreach (LatexDiagnostic diagnostic in parse.Diagnostics) {
            cancellationToken.ThrowIfCancellationRequested();
            int low = 0, high = blocks.Count;
            while (low < high) {
                int middle = low + (high - low) / 2;
                if (blocks[middle].Span.Start.Offset <= diagnostic.Span.Start.Offset) low = middle + 1;
                else high = middle;
            }
            int index = low - 1;
            if (index >= 0 && diagnostic.Span.End.Offset <= blocks[index].Span.End.Offset) byBlock[index].Add(diagnostic);
            else unattached.Add(diagnostic);
        }
        return unattached;
    }

    private static int InclusiveEnd(LatexSourceSpan span) => span.End.Column == 1 && span.End.Line > span.Start.Line ? span.End.Line - 1 : span.End.Line;

    private sealed class HeadingState {
        internal HeadingState(int level, string title) { Level = level; Title = title; }
        internal int Level { get; }
        internal string Title { get; }
    }
}
