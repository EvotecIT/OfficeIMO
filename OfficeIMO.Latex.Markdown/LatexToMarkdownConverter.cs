namespace OfficeIMO.Latex.Markdown;

/// <summary>Converts the bounded OfficeIMO LaTeX profile to typed Markdown.</summary>
internal static class LatexToMarkdownConverter {
    /// <summary>Converts recognized semantics and diagnoses source fallbacks.</summary>
    internal static LatexToMarkdownResult Convert(
        LatexDocument document,
        LatexToMarkdownOptions? options = null) {
        return Project(document, options).Result;
    }

    internal static LatexMarkdownProjection Project(LatexDocument document, LatexToMarkdownOptions? options = null,
        System.Threading.CancellationToken cancellationToken = default) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        cancellationToken.ThrowIfCancellationRequested();
        document = document.GetCurrentView(cancellationToken);
        options ??= new LatexToMarkdownOptions();
        var target = MarkdownDoc.Create();
        var diagnostics = new List<LatexMarkdownConversionDiagnostic>();
        var blocks = new List<LatexProjectedBlock>();
        AddFrontMatter(document, target, options, diagnostics);
        LatexCommand? titleCommand = document.Commands.FirstOrDefault(static command => command.Name == "title" && LatexSemanticBuilder.IsActiveSyntax(command.Syntax));
        LatexArgument? title = titleCommand?.GetRequiredArgument(0);
        if (title != null && document.Profile != LatexDocumentProfile.PreserveOnly && document.Body != null) {
            var titleDiagnostics = new List<LatexMarkdownConversionDiagnostic>();
            var titleDoc = MarkdownDoc.Create().Add(new HeadingBlock(1, LatexInlineToMarkdownConverter.Convert(document, title.ContentSpan, titleDiagnostics)));
            diagnostics.AddRange(titleDiagnostics);
            foreach (IMarkdownBlock block in titleDoc.Blocks) target.Add(block);
            blocks.Add(new LatexProjectedBlock(titleCommand!.Syntax.Span, "title", titleDoc, titleDiagnostics));
        }
        BlockCandidate[] candidates = BuildCandidates(document)
            .OrderBy(static candidate => candidate.Span.Start.Offset)
            .ThenByDescending(static candidate => candidate.Span.Length).ToArray();
        int consumedUntil = document.Body?.ContentSpan.Start.Offset ?? 0;
        foreach (BlockCandidate candidate in candidates) {
            cancellationToken.ThrowIfCancellationRequested();
            if (candidate.Span.Start.Offset < consumedUntil) continue;
            var projected = MarkdownDoc.Create();
            var blockDiagnostics = new List<LatexMarkdownConversionDiagnostic>();
            AddCandidate(document, projected, candidate, options, blockDiagnostics);
            foreach (IMarkdownBlock block in projected.Blocks) target.Add(block);
            diagnostics.AddRange(blockDiagnostics);
            blocks.Add(new LatexProjectedBlock(candidate.Span, candidate.Kind, projected, blockDiagnostics));
            consumedUntil = candidate.Span.End.Offset;
        }
        return new LatexMarkdownProjection(document, new LatexToMarkdownResult(target, diagnostics), blocks);
    }

    private static IEnumerable<BlockCandidate> BuildCandidates(LatexDocument document) {
        if (document.Body == null || document.Profile == LatexDocumentProfile.PreserveOnly) {
            LatexSourceSpan fallback = document.Body?.ContentSpan ?? document.SyntaxTree.Root.Span;
            yield return new BlockCandidate(fallback, fallback);
            yield break;
        }
        int start = document.Body.ContentSpan.Start.Offset;
        int end = document.Body.ContentSpan.End.Offset;
        foreach (LatexHeading heading in document.Headings.Where(heading => IsInside(heading.Command.Syntax.Span, start, end))) {
            yield return new BlockCandidate(heading.Command.Syntax.Span, heading);
        }
        foreach (LatexParagraph paragraph in document.Paragraphs) yield return new BlockCandidate(paragraph.Span, paragraph);
        foreach (LatexList list in document.Lists.Where(list => IsInside(list.Environment.Syntax.Span, start, end))) {
            yield return new BlockCandidate(list.Environment.Syntax.Span, list);
        }
        foreach (LatexFigure figure in document.Figures.Where(figure => IsInside(figure.Environment.Syntax.Span, start, end))) {
            yield return new BlockCandidate(figure.Environment.Syntax.Span, figure);
        }
        foreach (LatexTable table in document.Tables.Where(table => IsInside(table.Environment.Syntax.Span, start, end))) {
            LatexEnvironment? container = FindAncestorEnvironment(document, table.Environment, "table");
            yield return new BlockCandidate(container?.Syntax.Span ?? table.Environment.Syntax.Span, table);
        }
        foreach (LatexTheorem theorem in document.Theorems.Where(theorem => IsInside(theorem.Environment.Syntax.Span, start, end))) {
            yield return new BlockCandidate(theorem.Environment.Syntax.Span, theorem);
        }
        foreach (LatexMath math in document.Math.Where(math =>
                     math.Kind != LatexMathKind.InlineDollar && math.Kind != LatexMathKind.InlineParentheses &&
                     IsInside(math.Syntax.Span, start, end))) {
            yield return new BlockCandidate(math.Syntax.Span, math);
        }
        foreach (LatexSyntaxNode verbatim in document.SyntaxTree.Root.DescendantsAndSelf().Where(node =>
                     node.Kind == LatexSyntaxKind.Verbatim && LatexSemanticBuilder.IsActiveSyntax(node) && !LatexSemanticBuilder.IsInsideCommandArgument(node) &&
                     !string.Equals(node.Value, "verb", StringComparison.Ordinal) &&
                     IsInside(node.Span, start, end) && IsDirectChildSyntax(node, document.Body.Syntax))) {
            yield return new BlockCandidate(verbatim.Span, verbatim);
        }
        foreach (LatexEnvironment environment in document.Environments.Where(environment =>
                     !ReferenceEquals(environment, document.Body) && IsInside(environment.Syntax.Span, start, end) &&
                     IsDirectChildEnvironment(environment.Syntax, document.Body.Syntax) &&
                     !HasSemanticProjection(document, environment))) {
            yield return new BlockCandidate(environment.Syntax.Span, environment);
        }
    }

    private static void AddCandidate(
        LatexDocument document,
        MarkdownDoc target,
        BlockCandidate candidate,
        LatexToMarkdownOptions options,
        List<LatexMarkdownConversionDiagnostic> diagnostics) {
        switch (candidate.Value) {
            case LatexSourceSpan fallback:
                AddSourceFallback(document, target, fallback, options, diagnostics);
                break;
            case LatexHeading heading: {
                    LatexArgument title = heading.Command.GetRequiredArgument(0)!;
                    int markdownLevel = GetMarkdownHeadingLevel(document, heading);
                    var block = new HeadingBlock(markdownLevel,
                        LatexInlineToMarkdownConverter.Convert(document, title.ContentSpan, diagnostics));
                    ApplyLabel(document, block, candidate.Span, diagnostics);
                    target.Add(block);
                    diagnostics.Add(new LatexMarkdownConversionDiagnostic(
                        "LATEXMD212",
                        LatexMarkdownConversionOutcome.Simplified,
                        "heading-numbering",
                        heading.IsStarred
                            ? "The unnumbered LaTeX heading was mapped to a Markdown heading; starred numbering and table-of-contents behavior is not expressible in plain Markdown."
                            : "The numbered LaTeX heading was mapped to a Markdown heading; automatic numbering and table-of-contents behavior is not expressible in plain Markdown.",
                        heading.Command.Syntax.Span));
                    break;
                }
            case LatexParagraph paragraph: {
                    InlineSequence inlines = LatexInlineToMarkdownConverter.Convert(document, paragraph.Span, diagnostics);
                    if (inlines.Nodes.Count > 0) target.Add(new ParagraphBlock(inlines));
                    break;
                }
            case LatexList list:
                AddList(document, target, list, diagnostics);
                break;
            case LatexFigure figure:
                AddFigure(document, target, figure, options, diagnostics);
                break;
            case LatexTable table:
                target.Add(ConvertTable(document, table, diagnostics));
                LatexEnvironment? tableContainer = FindAncestorEnvironment(document, table.Environment, "table");
                if (tableContainer != null) {
                    var represented = new List<LatexSourceSpan> { table.Environment.Syntax.Span };
                    LatexCommand? caption = FindDirectCommand(document, tableContainer, "caption");
                    LatexCommand? label = FindDirectCommand(document, tableContainer, "label");
                    if (caption != null) represented.Add(caption.Syntax.Span);
                    if (label != null) represented.Add(label.Syntax.Span);
                    AddResidualSource(document, target, tableContainer, represented, options, diagnostics, "table-container");
                }
                break;
            case LatexTheorem theorem:
                AddTheorem(document, target, theorem, diagnostics);
                break;
            case LatexMath math:
                target.Add(new SemanticFencedBlock(MarkdownSemanticKinds.Math, "latex", ExtractVisibleSource(document, math.Syntax, math.ContentSpan, diagnostics)));
                diagnostics.Add(new LatexMarkdownConversionDiagnostic(
                    "LATEXMD201", LatexMarkdownConversionOutcome.Simplified, "display-math",
                    "Display math source was transported without TeX layout evaluation.", math.Syntax.Span));
                break;
            case LatexEnvironment environment:
                AddEnvironmentFallback(document, target, environment, options, diagnostics);
                break;
            case LatexSyntaxNode verbatim:
                AddVerbatimBlock(target, verbatim, diagnostics);
                break;
        }
    }

    private static void AddVerbatimBlock(
        MarkdownDoc target,
        LatexSyntaxNode syntax,
        List<LatexMarkdownConversionDiagnostic> diagnostics) {
        if (string.Equals(syntax.Value, "comment", StringComparison.Ordinal)) {
            diagnostics.Add(new LatexMarkdownConversionDiagnostic(
                "LATEXMD210", LatexMarkdownConversionOutcome.Omitted, "comment-environment",
                "The LaTeX comment environment was omitted and its body was not exposed as Markdown text.", syntax.Span));
            return;
        }
        target.Add(new CodeBlock("text", LatexInlineToMarkdownConverter.GetVerbatimContent(syntax)));
        diagnostics.Add(new LatexMarkdownConversionDiagnostic(
            "LATEXMD213", LatexMarkdownConversionOutcome.Simplified, "verbatim-environment",
            "Opaque LaTeX verbatim content was retained as a fenced code block without TeX environment semantics.", syntax.Span));
    }

    private static int GetMarkdownHeadingLevel(LatexDocument document, LatexHeading heading) {
        int firstSectionLevel = string.Equals(document.DocumentClassName, "article", StringComparison.Ordinal) ? 2 : 1;
        bool hasPart = document.Body != null && document.Headings.Any(candidate => candidate.Level == 0 &&
            IsInside(candidate.Command.Syntax.Span, document.Body.ContentSpan.Start.Offset, document.Body.ContentSpan.End.Offset));
        if (hasPart && heading.Level == 0) return 1;
        int markdownLevel = heading.Level - firstSectionLevel + 1;
        if (hasPart) markdownLevel++;
        return Math.Max(1, Math.Min(6, markdownLevel));
    }

    private static void AddList(
        LatexDocument document,
        MarkdownDoc target,
        LatexList source,
        List<LatexMarkdownConversionDiagnostic> diagnostics) {
        if (source.Kind == LatexListKind.Description) {
            var definitions = new DefinitionListBlock();
            foreach (LatexListItem item in source.Items) {
                var term = item.ItemCommand.GetOptionalArgument(0) is LatexArgument label
                    ? LatexInlineToMarkdownConverter.Convert(document, label.ContentSpan, diagnostics)
                    : new InlineSequence { AutoSpacing = false };
                var body = new ParagraphBlock(LatexInlineToMarkdownConverter.Convert(document, item.ContentSpan, diagnostics));
                definitions.AddEntry(new DefinitionListEntry(term, new IMarkdownBlock[] { body }));
            }
            target.Add(definitions);
            return;
        }
        if (source.Kind == LatexListKind.Ordered) {
            var list = new OrderedListBlock();
            foreach (LatexListItem item in source.Items) {
                list.Items.Add(new ListItem(ConvertListItem(document, item, diagnostics)));
            }
            target.Add(list);
        } else {
            var list = new UnorderedListBlock();
            foreach (LatexListItem item in source.Items) {
                list.Items.Add(new ListItem(ConvertListItem(document, item, diagnostics)));
            }
            target.Add(list);
        }
    }

    private static InlineSequence ConvertListItem(
        LatexDocument document,
        LatexListItem item,
        List<LatexMarkdownConversionDiagnostic> diagnostics) {
        InlineSequence content = LatexInlineToMarkdownConverter.Convert(document, item.ContentSpan, diagnostics);
        if (item.ItemCommand.GetOptionalArgument(0) is LatexArgument label) {
            InlineSequence labeled = LatexInlineToMarkdownConverter.Convert(document, label.ContentSpan, diagnostics);
            labeled.AddRaw(new MarkdownTextRun(": "));
            foreach (IMarkdownInline inline in content.Nodes) labeled.AddRaw(inline);
            diagnostics.Add(new LatexMarkdownConversionDiagnostic(
                "LATEXMD214", LatexMarkdownConversionOutcome.Simplified, "list-item-label",
                "The custom TeX item label was retained as visible item text; target list markers use Markdown numbering or bullets.", item.ItemCommand.Syntax.Span));
            return labeled;
        }
        return content;
    }

    private static void AddFigure(
        LatexDocument document,
        MarkdownDoc target,
        LatexFigure source,
        LatexToMarkdownOptions options,
        List<LatexMarkdownConversionDiagnostic> diagnostics) {
        for (int index = 0; index < source.Images.Count; index++) {
            LatexImage image = source.Images[index];
            LatexInlineToMarkdownConverter.ReportGraphicsOptions(image.Command, diagnostics);
            string caption = LatexInlineToMarkdownConverter.ReadArgumentSource(document, source.CaptionCommand?.GetRequiredArgument(0), diagnostics);
            string label = LatexInlineToMarkdownConverter.ReadArgumentSource(document, source.LabelCommand?.GetRequiredArgument(0), diagnostics);
            var block = new ImageBlock(LatexLiteralText.Decode(LatexInlineToMarkdownConverter.ReadArgumentSource(document, image.Command.GetRequiredArgument(0), diagnostics)), caption);
            if (!string.IsNullOrWhiteSpace(label)) block.SetAttributes(MarkdownAttributeSet.Create(label));
            if (!string.IsNullOrWhiteSpace(caption)) block.Caption = caption;
            target.Add(block);
        }
        if (source.Images.Count == 0) {
            AddEnvironmentFallback(document, target, source.Environment, options, diagnostics);
            return;
        }
        var represented = source.Images.Select(static image => image.Command.Syntax.Span).ToList();
        if (source.CaptionCommand != null) represented.Add(source.CaptionCommand.Syntax.Span);
        if (source.LabelCommand != null) represented.Add(source.LabelCommand.Syntax.Span);
        AddResidualSource(document, target, source.Environment, represented, options, diagnostics, "figure-content");
    }

    private static TableBlock ConvertTable(
        LatexDocument document,
        LatexTable source,
        List<LatexMarkdownConversionDiagnostic> diagnostics) {
        var target = new TableBlock();
        bool header = source.Rows.Count > 1 && source.Rows[0].Cells.All(static cell => cell.Content.TrimStart().StartsWith("\\textbf", StringComparison.Ordinal));
        var structuredHeaders = new List<TableCell>();
        var structuredRows = new List<IReadOnlyList<TableCell>>();
        if (!header && source.Rows.Count > 0) {
            int columnCount = Math.Max(1, source.Rows.Max(static row => row.Cells.Count));
            for (int index = 0; index < columnCount; index++) {
                target.Headers.Add(string.Empty);
                structuredHeaders.Add(new TableCell(new IMarkdownBlock[] { new ParagraphBlock(new InlineSequence()) }));
            }
            diagnostics.Add(new LatexMarkdownConversionDiagnostic(
                "LATEXMD211", LatexMarkdownConversionOutcome.Simplified, "table-header",
                "A blank Markdown header was added because pipe tables require a header row.", source.Environment.Syntax.Span));
        }
        for (int rowIndex = 0; rowIndex < source.Rows.Count; rowIndex++) {
            var values = new List<string>();
            var cells = new List<TableCell>();
            foreach (LatexTableCell sourceCell in source.Rows[rowIndex].Cells) {
                InlineSequence inlines = LatexInlineToMarkdownConverter.Convert(document, sourceCell.Span, diagnostics);
                string markdown = MarkdownDoc.Create().Add(new ParagraphBlock(inlines)).ToMarkdown().Trim();
                values.Add(markdown);
                cells.Add(new TableCell(new IMarkdownBlock[] { new ParagraphBlock(inlines) }));
            }
            if (header && rowIndex == 0) {
                target.Headers.AddRange(values);
                structuredHeaders.AddRange(cells);
            } else {
                target.Rows.Add(values);
                structuredRows.Add(cells);
            }
        }
        target.SetStructuredCells(structuredHeaders, structuredRows, target.ComputeContentSignature());
        LatexEnvironment? container = FindAncestorEnvironment(document, source.Environment, "table");
        LatexCommand? caption = FindDirectCommand(document, container, "caption");
        LatexCommand? label = FindDirectCommand(document, container, "label");
        string captionText = LatexInlineToMarkdownConverter.ReadArgumentSource(document, caption?.GetRequiredArgument(0), diagnostics);
        string labelText = LatexInlineToMarkdownConverter.ReadArgumentSource(document, label?.GetRequiredArgument(0), diagnostics);
        if (!string.IsNullOrWhiteSpace(captionText) || !string.IsNullOrWhiteSpace(labelText)) {
            var attributes = string.IsNullOrWhiteSpace(captionText)
                ? null
                : new[] { new KeyValuePair<string, string?>("caption", captionText) };
            target.SetAttributes(MarkdownAttributeSet.Create(labelText, attributes: attributes));
        }
        return target;
    }

    private static void AddTheorem(
        LatexDocument document,
        MarkdownDoc target,
        LatexTheorem source,
        List<LatexMarkdownConversionDiagnostic> diagnostics) {
        InlineSequence body = source.LabelCommand == null
            ? LatexInlineToMarkdownConverter.Convert(document, source.Environment.ContentSpan, diagnostics)
            : LatexInlineToMarkdownConverter.ConvertExcluding(
                document,
                source.Environment.ContentSpan,
                new[] { source.LabelCommand.Syntax.Span },
                diagnostics);
        var callout = new CalloutBlock(source.Kind, LatexLiteralText.Decode(LatexInlineToMarkdownConverter.ReadArgumentSource(document, source.Environment.BeginCommand.GetOptionalArgument(0), diagnostics)),
            new IMarkdownBlock[] { new ParagraphBlock(body) });
        string label = LatexInlineToMarkdownConverter.ReadArgumentSource(document, source.LabelCommand?.GetRequiredArgument(0), diagnostics);
        if (!string.IsNullOrWhiteSpace(label)) callout.SetAttributes(MarkdownAttributeSet.Create(label));
        target.Add(callout);
    }

    private static void AddEnvironmentFallback(
        LatexDocument document,
        MarkdownDoc target,
        LatexEnvironment source,
        LatexToMarkdownOptions options,
        List<LatexMarkdownConversionDiagnostic> diagnostics) {
        if (string.Equals(source.Name, "quote", StringComparison.Ordinal) || string.Equals(source.Name, "quotation", StringComparison.Ordinal)) {
            var quote = new QuoteBlock();
            quote.ChildBlocks.Add(new ParagraphBlock(LatexInlineToMarkdownConverter.Convert(document, source.ContentSpan, diagnostics)));
            target.Add(quote);
            return;
        }
        if (string.Equals(source.Name, "verbatim", StringComparison.Ordinal)) {
            target.Code("text", source.Content.Trim('\r', '\n'));
            return;
        }
        LatexSyntaxNode[] comments = FindCommentEnvironments(source.Syntax).ToArray();
        if (options.PreserveUnsupportedAsSource) {
            string visibleSource = ExtractResidual(document.Source.Text, source.Syntax.Span,
                comments.Select(static comment => comment.Span).Concat(document.Tokens.Where(static token => token.Kind == LatexTokenKind.Comment).Select(static token => token.Span)));
            if (!string.IsNullOrWhiteSpace(visibleSource)) target.Code("latex", visibleSource);
        }
        ReportOmittedComments(comments, diagnostics);
        diagnostics.Add(new LatexMarkdownConversionDiagnostic(
            "LATEXMD299",
            options.PreserveUnsupportedAsSource ? LatexMarkdownConversionOutcome.SourceFallback : LatexMarkdownConversionOutcome.Omitted,
            "environment:" + source.Name,
            options.PreserveUnsupportedAsSource ? "Unknown environment retained as visible LaTeX source." : "Unknown environment omitted by conversion options.",
            source.Syntax.Span));
    }

    private static void AddResidualSource(
        LatexDocument document,
        MarkdownDoc target,
        LatexEnvironment environment,
        IEnumerable<LatexSourceSpan> representedSpans,
        LatexToMarkdownOptions options,
        List<LatexMarkdownConversionDiagnostic> diagnostics,
        string feature) {
        LatexSyntaxNode[] comments = FindCommentEnvironments(environment.Syntax).ToArray();
        string residual = ExtractResidual(document.Source.Text, environment.ContentSpan,
            representedSpans.Concat(comments.Select(static comment => comment.Span)).Concat(document.Tokens.Where(static token => token.Kind == LatexTokenKind.Comment).Select(static token => token.Span)));
        ReportOmittedComments(comments, diagnostics);
        if (string.IsNullOrWhiteSpace(residual)) return;
        if (options.PreserveUnsupportedAsSource) target.Code("latex", residual.Trim());
        diagnostics.Add(new LatexMarkdownConversionDiagnostic(
            "LATEXMD298",
            options.PreserveUnsupportedAsSource ? LatexMarkdownConversionOutcome.SourceFallback : LatexMarkdownConversionOutcome.Omitted,
            feature,
            options.PreserveUnsupportedAsSource
                ? "Unrepresented environment content was retained as visible LaTeX source."
                : "Unrepresented environment content was omitted by conversion options.",
            environment.Syntax.Span));
    }

    private static IEnumerable<LatexSyntaxNode> FindCommentEnvironments(LatexSyntaxNode syntax) =>
        syntax.DescendantsAndSelf().Where(static node =>
            node.Kind == LatexSyntaxKind.Verbatim
            && string.Equals(node.Value, "comment", StringComparison.Ordinal));

    private static void ReportOmittedComments(
        IEnumerable<LatexSyntaxNode> comments,
        List<LatexMarkdownConversionDiagnostic> diagnostics) {
        foreach (LatexSyntaxNode comment in comments) {
            diagnostics.Add(new LatexMarkdownConversionDiagnostic(
                "LATEXMD210", LatexMarkdownConversionOutcome.Omitted, "comment-environment",
                "The LaTeX comment environment was omitted and its body was not exposed as Markdown text.",
                comment.Span));
        }
    }

    internal static string ExtractVisibleSource(LatexDocument document, LatexSyntaxNode syntax, LatexSourceSpan span,
        List<LatexMarkdownConversionDiagnostic> diagnostics) {
        LatexSyntaxNode[] comments = syntax.DescendantsAndSelf().Where(static node => node.Kind == LatexSyntaxKind.Comment ||
            (node.Kind == LatexSyntaxKind.Verbatim && node.Value == "comment")).ToArray();
        ReportOmittedComments(comments.Where(static node => node.Kind == LatexSyntaxKind.Verbatim), diagnostics);
        return ExtractResidual(document.Source.Text, span, comments.Select(static node => node.Span));
    }

    private static string ExtractResidual(
        string source,
        LatexSourceSpan contentSpan,
        IEnumerable<LatexSourceSpan> representedSpans) {
        var output = new StringBuilder();
        int cursor = contentSpan.Start.Offset;
        foreach (LatexSourceSpan represented in representedSpans
                     .Where(span => span.End.Offset > contentSpan.Start.Offset && span.Start.Offset < contentSpan.End.Offset)
                     .OrderBy(static span => span.Start.Offset)) {
            int start = Math.Max(cursor, represented.Start.Offset);
            int end = Math.Min(contentSpan.End.Offset, represented.End.Offset);
            if (start > cursor) output.Append(source, cursor, start - cursor);
            cursor = Math.Max(cursor, end);
        }
        if (cursor < contentSpan.End.Offset) output.Append(source, cursor, contentSpan.End.Offset - cursor);
        return output.ToString();
    }

    private static void AddFrontMatter(LatexDocument source, MarkdownDoc target, LatexToMarkdownOptions options,
        List<LatexMarkdownConversionDiagnostic> diagnostics) {
        if (!options.IncludePreambleAsFrontMatter) return;
        var values = new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase);
        if (source.DocumentClassName != null) values["documentclass"] = source.DocumentClassName;
        AddCommandValue(source, values, "title", diagnostics);
        AddCommandValue(source, values, "author", diagnostics);
        AddCommandValue(source, values, "date", diagnostics);
        if (values.Count > 0) target.FrontMatter(values);
    }

    private static void AddCommandValue(LatexDocument source, Dictionary<string, object?> values, string name,
        List<LatexMarkdownConversionDiagnostic> diagnostics) {
        LatexArgument? argument = source.Commands.FirstOrDefault(command => string.Equals(command.Name, name, StringComparison.Ordinal) && LatexSemanticBuilder.IsActiveSyntax(command.Syntax))?.GetRequiredArgument(0);
        string value = LatexInlineToMarkdownConverter.ReadArgumentSource(source, argument, diagnostics);
        if (!string.IsNullOrEmpty(value)) values[name] = value;
    }

    private static void ApplyLabel(LatexDocument document, MarkdownObject target, LatexSourceSpan owner,
        List<LatexMarkdownConversionDiagnostic> diagnostics) {
        LatexLabel? label = FindAdjacentLabel(document, owner);
        if (label != null) target.SetAttributes(MarkdownAttributeSet.Create(LatexInlineToMarkdownConverter.ReadArgumentSource(document, label.Command.GetRequiredArgument(0), diagnostics)));
    }

    private static LatexLabel? FindAdjacentLabel(LatexDocument document, LatexSourceSpan owner) =>
        document.Labels.FirstOrDefault(item =>
            item.Command.Syntax.Span.Start.Offset >= owner.End.Offset &&
            IsWhitespaceOnly(document.Source.Text, owner.End.Offset, item.Command.Syntax.Span.Start.Offset));

    private static LatexCommand? FindDirectCommand(LatexDocument document, LatexEnvironment? environment, string name) {
        if (environment == null) return null;
        return document.Commands.FirstOrDefault(command => string.Equals(command.Name, name, StringComparison.Ordinal) &&
            IsDirectlyInside(command.Syntax, environment.Syntax));
    }

    private static bool IsDirectlyInside(LatexSyntaxNode node, LatexSyntaxNode environment) {
        LatexSyntaxNode? current = node.Parent;
        while (current != null) {
            if (current.Kind == LatexSyntaxKind.Environment) return ReferenceEquals(current, environment);
            current = current.Parent;
        }
        return false;
    }

    private static bool IsWhitespaceOnly(string source, int start, int end) {
        for (int index = start; index < end; index++) {
            if (!char.IsWhiteSpace(source[index])) return false;
        }
        return true;
    }

    private static LatexEnvironment? FindAncestorEnvironment(LatexDocument document, LatexEnvironment source, string name) {
        LatexSyntaxNode? current = source.Syntax.Parent;
        while (current != null) {
            if (current.Kind == LatexSyntaxKind.Environment && string.Equals(current.Value, name, StringComparison.Ordinal)) {
                return document.Environments.FirstOrDefault(environment => ReferenceEquals(environment.Syntax, current));
            }
            current = current.Parent;
        }
        return null;
    }

    private static bool HasSemanticProjection(LatexDocument document, LatexEnvironment environment) =>
        document.Lists.Any(item => ReferenceEquals(item.Environment, environment)) ||
        document.Figures.Any(item => ReferenceEquals(item.Environment, environment)) ||
        document.Tables.Any(item => ReferenceEquals(item.Environment, environment) || ReferenceEquals(FindAncestorEnvironment(document, item.Environment, "table"), environment)) ||
        document.Theorems.Any(item => ReferenceEquals(item.Environment, environment)) ||
        document.Math.Any(item => ReferenceEquals(item.Environment, environment));

    private static void AddSourceFallback(LatexDocument document, MarkdownDoc target, LatexSourceSpan span,
        LatexToMarkdownOptions options, List<LatexMarkdownConversionDiagnostic> diagnostics) {
        LatexSyntaxNode[] comments = FindCommentEnvironments(document.SyntaxTree.Root).ToArray();
        string visible = ExtractResidual(document.Source.Text, span,
            comments.Select(static item => item.Span).Concat(document.Tokens.Where(static token => token.Kind == LatexTokenKind.Comment).Select(static token => token.Span)));
        ReportOmittedComments(comments.Where(item => IsInside(item.Span, span.Start.Offset, span.End.Offset)), diagnostics);
        if (string.IsNullOrWhiteSpace(visible)) return;
        if (options.PreserveUnsupportedAsSource) target.Code("latex", visible.Trim());
        diagnostics.Add(new LatexMarkdownConversionDiagnostic("LATEXMD297",
            options.PreserveUnsupportedAsSource ? LatexMarkdownConversionOutcome.SourceFallback : LatexMarkdownConversionOutcome.Omitted,
            "unprojected-source", "Source outside the semantic document profile requires a source fallback.", span));
    }

    private static bool IsDirectChildEnvironment(LatexSyntaxNode node, LatexSyntaxNode body) {
        LatexSyntaxNode? current = node.Parent;
        while (current != null) {
            if (current.Kind == LatexSyntaxKind.Environment) return ReferenceEquals(current, body);
            current = current.Parent;
        }
        return false;
    }

    private static bool IsDirectChildSyntax(LatexSyntaxNode node, LatexSyntaxNode body) {
        LatexSyntaxNode? current = node.Parent;
        while (current != null) {
            if (current.Kind == LatexSyntaxKind.Environment) return ReferenceEquals(current, body);
            current = current.Parent;
        }
        return false;
    }

    private static bool IsInside(LatexSourceSpan span, int start, int end) => span.Start.Offset >= start && span.End.Offset <= end;

    private sealed class BlockCandidate {
        internal BlockCandidate(LatexSourceSpan span, object value) { Span = span; Value = value; }
        internal LatexSourceSpan Span { get; }
        internal object Value { get; }
        internal string Kind => Value switch {
            LatexHeading => "heading",
            LatexParagraph => "paragraph",
            LatexList list => "list-" + list.Kind.ToString().ToLowerInvariant(),
            LatexFigure => "figure",
            LatexTable => "table",
            LatexTheorem theorem => "theorem-" + theorem.Kind,
            LatexMath => "math-display",
            LatexEnvironment environment => "environment-" + environment.Name,
            LatexSyntaxNode => "verbatim",
            _ => "source-fallback"
        };
    }
}

/// <summary>Conversion extensions.</summary>
public static partial class LatexMarkdownConverterExtensions {
    /// <summary>Converts a native LaTeX document to Markdown.</summary>
    public static LatexToMarkdownResult ToMarkdownDocumentResult(this LatexDocument document, LatexToMarkdownOptions? options = null) =>
        LatexToMarkdownConverter.Convert(document, options);

    /// <summary>Converts the current document with cooperative cancellation, rebinding edited source before projection.</summary>
    public static LatexToMarkdownResult ToMarkdownDocumentResult(this LatexDocument document, LatexToMarkdownOptions? options,
        System.Threading.CancellationToken cancellationToken) => LatexToMarkdownConverter.Project(document, options, cancellationToken).Result;

    /// <summary>Converts a LaTeX document to a typed Markdown document.</summary>
    public static MarkdownDoc ToMarkdownDocument(this LatexDocument document, LatexToMarkdownOptions? options = null) =>
        document.ToMarkdownDocumentResult(options).Value;
}
