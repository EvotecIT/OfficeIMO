namespace OfficeIMO.Reader.Latex;

internal static class LatexReaderChunkBuilder {
    internal static IEnumerable<ReaderChunk> BuildBlocks(
        LatexParseResult result,
        string sourceName,
        ReaderOptions readerOptions,
        ReaderLatexOptions options,
        CancellationToken cancellationToken) {
        var headingStack = new List<HeadingState>();
        Candidate[] candidates = BuildCandidates(result.Document).OrderBy(static candidate => candidate.Span.Start.Offset)
            .ThenByDescending(static candidate => candidate.Span.Length).ToArray();
        if (candidates.Length == 0 && result.Document.Source.Text.Length > 0) {
            candidates = new[] { new Candidate(result.Document.SyntaxTree.Root.Span, result.Document.SyntaxTree.Root, "source-fallback") };
        }
        int consumedUntil = -1;
        int emitted = 0;
        for (int sourceIndex = 0; sourceIndex < candidates.Length; sourceIndex++) {
            cancellationToken.ThrowIfCancellationRequested();
            Candidate candidate = candidates[sourceIndex];
            if (candidate.Span.Start.Offset < consumedUntil) continue;
            ProjectedChunk projected = Project(result.Document, candidate, options);
            consumedUntil = candidate.Span.End.Offset;
            if (candidate.Value is LatexHeading heading) UpdateHeadingStack(headingStack, heading);
            IReadOnlyList<string> parts = Split(projected.Text.Length == 0 ? projected.Markdown : projected.Text, readerOptions.MaxChars);
            if (parts.Count == 0) continue;
            for (int partIndex = 0; partIndex < parts.Count; partIndex++) {
                yield return new ReaderChunk {
                    Id = parts.Count == 1 ? "latex-" + sourceIndex : "latex-" + sourceIndex + "-part-" + (partIndex + 1),
                    Kind = ReaderInputKind.Latex,
                    Location = new ReaderLocation {
                        Path = sourceName,
                        BlockIndex = emitted++,
                        SourceBlockIndex = sourceIndex,
                        StartLine = candidate.Span.Start.Line,
                        EndLine = InclusiveEnd(candidate.Span),
                        HeadingPath = headingStack.Count == 0 ? null : string.Join(" > ", headingStack.Select(static item => item.Title)),
                        SourceBlockKind = candidate.Kind,
                        BlockAnchor = "latex-block-" + sourceIndex
                    },
                    Text = parts[partIndex],
                    Markdown = parts.Count == 1 ? projected.Markdown : parts[partIndex],
                    Diagnostics = new ReaderChunkDiagnostics { SourceKind = "latex" },
                    Warnings = options.IncludeDiagnostics ? BuildWarnings(result, projected.Diagnostics, candidate.Span, parts.Count > 1) : null
                };
            }
        }
    }

    internal static IEnumerable<ReaderChunk> BuildDocument(
        LatexParseResult result,
        string sourceName,
        ReaderOptions readerOptions,
        ReaderLatexOptions options,
        CancellationToken cancellationToken) {
        LatexMarkdownConversionResult conversion = result.Document.ToMarkdownDocument(options.MarkdownOptions);
        string markdown = conversion.Document.ToMarkdown().TrimEnd();
        var projectedText = new List<string>();
        Candidate[] candidates = BuildCandidates(result.Document)
            .OrderBy(static candidate => candidate.Span.Start.Offset)
            .ThenByDescending(static candidate => candidate.Span.Length)
            .ToArray();
        if (candidates.Length == 0 && result.Document.Source.Text.Length > 0) {
            candidates = new[] { new Candidate(result.Document.SyntaxTree.Root.Span, result.Document.SyntaxTree.Root, "source-fallback") };
        }
        int consumedUntil = -1;
        for (int index = 0; index < candidates.Length; index++) {
            Candidate candidate = candidates[index];
            if (candidate.Span.Start.Offset < consumedUntil) continue;
            ProjectedChunk projected = Project(result.Document, candidate, options);
            if (!string.IsNullOrWhiteSpace(projected.Text)) projectedText.Add(projected.Text);
            if (markdown.Length == 0 && projected.Markdown.Length > 0) markdown = projected.Markdown;
            consumedUntil = candidate.Span.End.Offset;
        }
        string text = string.Join("\n\n", projectedText);
        IReadOnlyList<string> parts = Split(text.Length == 0 ? markdown : text, readerOptions.MaxChars);
        for (int index = 0; index < parts.Count; index++) {
            cancellationToken.ThrowIfCancellationRequested();
            yield return new ReaderChunk {
                Id = parts.Count == 1 ? "latex-document" : "latex-document-part-" + (index + 1),
                Kind = ReaderInputKind.Latex,
                Location = new ReaderLocation {
                    Path = sourceName,
                    BlockIndex = index,
                    SourceBlockIndex = 0,
                    StartLine = 1,
                    EndLine = result.Document.Source.LineCount,
                    SourceBlockKind = "document",
                    BlockAnchor = "latex-document"
                },
                Text = parts[index],
                Markdown = parts.Count == 1 ? markdown : parts[index],
                Diagnostics = new ReaderChunkDiagnostics { SourceKind = "latex" },
                Warnings = options.IncludeDiagnostics ? BuildWarnings(result, conversion.Diagnostics, result.Document.SyntaxTree.Root.Span, parts.Count > 1) : null
            };
        }
    }

    private static IEnumerable<Candidate> BuildCandidates(LatexDocument document) {
        LatexCommand? title = document.Commands.FirstOrDefault(static command => command.Name == "title");
        if (title?.GetRequiredArgument(0) != null) yield return new Candidate(title.Syntax.Span, title, "title");
        foreach (LatexHeading heading in document.Headings) yield return new Candidate(heading.Command.Syntax.Span, heading, "heading");
        foreach (LatexParagraph paragraph in document.Paragraphs) yield return new Candidate(paragraph.Span, paragraph, "paragraph");
        foreach (LatexList list in document.Lists) yield return new Candidate(list.Environment.Syntax.Span, list, "list-" + list.Kind.ToString().ToLowerInvariant());
        foreach (LatexFigure figure in document.Figures) yield return new Candidate(figure.Environment.Syntax.Span, figure, "figure");
        foreach (LatexTable table in document.Tables) {
            LatexEnvironment? container = FindTableContainer(document, table);
            yield return new Candidate(container?.Syntax.Span ?? table.Environment.Syntax.Span, table, "table");
        }
        foreach (LatexTheorem theorem in document.Theorems) yield return new Candidate(theorem.Environment.Syntax.Span, theorem, "theorem-" + theorem.Kind);
        foreach (LatexMath math in document.Math.Where(static math => math.Kind != LatexMathKind.InlineDollar && math.Kind != LatexMathKind.InlineParentheses)) {
            yield return new Candidate(math.Syntax.Span, math, "math-display");
        }
    }

    private static ProjectedChunk Project(LatexDocument document, Candidate candidate, ReaderLatexOptions options) {
        var diagnostics = new List<LatexMarkdownConversionDiagnostic>();
        var markdown = MarkdownDoc.Create();
        string text;
        switch (candidate.Value) {
            case LatexCommand titleCommand: {
                LatexArgument title = titleCommand.GetRequiredArgument(0)!;
                InlineSequence inlines = LatexInlineToMarkdownConverter.Convert(document, title.ContentSpan, diagnostics);
                markdown.Add(new HeadingBlock(1, inlines));
                text = InlinePlainText.Extract(inlines);
                break;
            }
            case LatexHeading heading: {
                InlineSequence inlines = LatexInlineToMarkdownConverter.Convert(document, heading.Command.GetRequiredArgument(0)!.ContentSpan, diagnostics);
                markdown.Add(new HeadingBlock(Math.Max(1, Math.Min(6, heading.Level)), inlines));
                text = InlinePlainText.Extract(inlines);
                break;
            }
            case LatexParagraph paragraph: {
                InlineSequence inlines = LatexInlineToMarkdownConverter.Convert(document, paragraph.Span, diagnostics);
                markdown.Add(new ParagraphBlock(inlines));
                text = InlinePlainText.Extract(inlines);
                break;
            }
            case LatexList list:
                text = ProjectList(document, list, markdown, diagnostics);
                break;
            case LatexFigure figure:
                var figureText = new List<string>();
                if (!string.IsNullOrWhiteSpace(figure.Caption)) figureText.Add(figure.Caption!);
                figureText.AddRange(figure.Images.Select(static image => image.Target));
                text = string.Join("\n", figureText);
                foreach (LatexImage image in figure.Images) markdown.Image(image.Target, figure.Caption);
                break;
            case LatexTable table:
                text = ProjectTable(document, table, markdown);
                break;
            case LatexTheorem theorem: {
                InlineSequence inlines = theorem.LabelCommand == null
                    ? LatexInlineToMarkdownConverter.Convert(document, theorem.Environment.ContentSpan, diagnostics)
                    : LatexInlineToMarkdownConverter.ConvertExcluding(
                        document,
                        theorem.Environment.ContentSpan,
                        new[] { theorem.LabelCommand.Syntax.Span },
                        diagnostics);
                markdown.Add(new CalloutBlock(theorem.Kind, theorem.Title ?? string.Empty, new IMarkdownBlock[] { new ParagraphBlock(inlines) }));
                text = InlinePlainText.Extract(inlines);
                break;
            }
            case LatexMath math:
                markdown.Add(new SemanticFencedBlock(MarkdownSemanticKinds.Math, "latex", math.Content));
                text = math.Content;
                break;
            default:
                text = candidate.Span.Slice(document.Source.Text);
                markdown.Code("latex", text);
                break;
        }
        return new ProjectedChunk(text, markdown.ToMarkdown().TrimEnd(), diagnostics);
    }

    private static string ProjectList(
        LatexDocument document,
        LatexList list,
        MarkdownDoc markdown,
        List<LatexMarkdownConversionDiagnostic> diagnostics) {
        var text = new List<string>();
        if (list.Kind == LatexListKind.Description) {
            var target = new DefinitionListBlock();
            foreach (LatexListItem item in list.Items) {
                InlineSequence inlines = LatexInlineToMarkdownConverter.Convert(document, item.ContentSpan, diagnostics);
                var term = new InlineSequence { AutoSpacing = false }.Text(item.Label ?? string.Empty);
                target.AddEntry(new DefinitionListEntry(term, new IMarkdownBlock[] { new ParagraphBlock(inlines) }));
                text.Add((item.Label == null ? string.Empty : item.Label + ": ") + InlinePlainText.Extract(inlines));
            }
            markdown.Add(target);
        } else if (list.Kind == LatexListKind.Ordered) {
            var target = new OrderedListBlock();
            foreach (LatexListItem item in list.Items) {
                InlineSequence inlines = LatexInlineToMarkdownConverter.Convert(document, item.ContentSpan, diagnostics);
                target.Items.Add(new ListItem(inlines));
                text.Add(InlinePlainText.Extract(inlines));
            }
            markdown.Add(target);
        } else {
            var target = new UnorderedListBlock();
            foreach (LatexListItem item in list.Items) {
                InlineSequence inlines = LatexInlineToMarkdownConverter.Convert(document, item.ContentSpan, diagnostics);
                target.Items.Add(new ListItem(inlines));
                text.Add((item.Label == null ? string.Empty : item.Label + ": ") + InlinePlainText.Extract(inlines));
            }
            markdown.Add(target);
        }
        return string.Join("\n", text);
    }

    private static string ProjectTable(LatexDocument document, LatexTable table, MarkdownDoc markdown) {
        var target = new TableBlock();
        int columns = table.Rows.Count == 0 ? 0 : table.Rows.Max(static row => row.Cells.Count);
        for (int index = 0; index < columns; index++) target.Headers.Add(string.Empty);
        foreach (LatexTableRow row in table.Rows) target.Rows.Add(row.Cells.Select(static cell => cell.Content).ToArray());
        LatexEnvironment? container = FindTableContainer(document, table);
        string? caption = FindContainerCommand(document, container, "caption")?.GetRequiredArgument(0)?.Content;
        string? label = FindContainerCommand(document, container, "label")?.GetRequiredArgument(0)?.Content;
        if (!string.IsNullOrWhiteSpace(caption) || !string.IsNullOrWhiteSpace(label)) {
            var attributes = string.IsNullOrWhiteSpace(caption)
                ? null
                : new[] { new KeyValuePair<string, string?>("caption", caption) };
            target.SetAttributes(MarkdownAttributeSet.Create(label, attributes: attributes));
        }
        markdown.Add(target);
        string rows = string.Join("\n", table.Rows.Select(row => string.Join("\t", row.Cells.Select(static cell => cell.Content))));
        return string.IsNullOrWhiteSpace(caption) ? rows : caption + (rows.Length == 0 ? string.Empty : "\n" + rows);
    }

    private static LatexEnvironment? FindTableContainer(LatexDocument document, LatexTable table) {
        for (LatexSyntaxNode? syntax = table.Environment.Syntax.Parent; syntax != null; syntax = syntax.Parent) {
            if (syntax.Kind == LatexSyntaxKind.Environment && string.Equals(syntax.Value, "table", StringComparison.Ordinal)) {
                return document.Environments.FirstOrDefault(environment => ReferenceEquals(environment.Syntax, syntax));
            }
        }
        return null;
    }

    private static LatexCommand? FindContainerCommand(LatexDocument document, LatexEnvironment? container, string name) {
        if (container == null) return null;
        return document.Commands.FirstOrDefault(command =>
            string.Equals(command.Name, name, StringComparison.Ordinal) &&
            command.Syntax.Span.Start.Offset >= container.ContentSpan.Start.Offset &&
            command.Syntax.Span.End.Offset <= container.ContentSpan.End.Offset);
    }

    private static IReadOnlyList<string>? BuildWarnings(
        LatexParseResult parse,
        IReadOnlyList<LatexMarkdownConversionDiagnostic> conversion,
        LatexSourceSpan span,
        bool split) {
        var warnings = parse.Diagnostics.Where(diagnostic => diagnostic.Span.Start.Offset >= span.Start.Offset && diagnostic.Span.End.Offset <= span.End.Offset)
            .Select(diagnostic => diagnostic.Code + ": " + diagnostic.Message).ToList();
        if (!parse.Document.IsRecognizedProfile) warnings.Add("LATEXR001: Source is not a recognized OfficeIMO LaTeX article, report, or book profile; preserved structures may be incomplete.");
        warnings.AddRange(conversion.Select(diagnostic => diagnostic.Code + ": " + diagnostic.Message));
        if (split) warnings.Add("LaTeX content was split due to ReaderOptions.MaxChars.");
        return warnings.Count == 0 ? null : warnings;
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

    private static void UpdateHeadingStack(List<HeadingState> stack, LatexHeading heading) {
        while (stack.Count > 0 && stack[stack.Count - 1].Level >= heading.Level) stack.RemoveAt(stack.Count - 1);
        stack.Add(new HeadingState(heading.Level, heading.Title));
    }

    private static int InclusiveEnd(LatexSourceSpan span) => span.End.Column == 1 && span.End.Line > span.Start.Line ? span.End.Line - 1 : span.End.Line;

    private sealed class Candidate {
        internal Candidate(LatexSourceSpan span, object value, string kind) { Span = span; Value = value; Kind = kind; }
        internal LatexSourceSpan Span { get; }
        internal object Value { get; }
        internal string Kind { get; }
    }

    private sealed class ProjectedChunk {
        internal ProjectedChunk(string text, string markdown, IReadOnlyList<LatexMarkdownConversionDiagnostic> diagnostics) {
            Text = text; Markdown = markdown; Diagnostics = diagnostics;
        }
        internal string Text { get; }
        internal string Markdown { get; }
        internal IReadOnlyList<LatexMarkdownConversionDiagnostic> Diagnostics { get; }
    }

    private sealed class HeadingState {
        internal HeadingState(int level, string title) { Level = level; Title = title; }
        internal int Level { get; }
        internal string Title { get; }
    }
}
