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
        var context = new LatexProjectionContext(document, cancellationToken);
        options ??= new LatexToMarkdownOptions();
        var target = MarkdownDoc.Create();
        var diagnostics = new List<LatexMarkdownConversionDiagnostic>();
        var blocks = new List<LatexProjectedBlock>();
        var globalDiagnostics = new List<LatexMarkdownConversionDiagnostic>();
        var aggregateBlocks = new List<IMarkdownBlock>();
        AddFrontMatter(context, target, options, globalDiagnostics);
        diagnostics.AddRange(globalDiagnostics);
        LatexCommand? titleCommand = context.FirstCommand("title");
        LatexArgument? title = titleCommand?.GetRequiredArgument(0);
        if (title != null && document.Profile != LatexDocumentProfile.PreserveOnly && document.Body != null) {
            var titleDiagnostics = new List<LatexMarkdownConversionDiagnostic>();
            var titleDoc = MarkdownDoc.Create().Add(new HeadingBlock(1, LatexInlineToMarkdownConverter.Convert(context, title.ContentSpan, titleDiagnostics)));
            diagnostics.AddRange(titleDiagnostics);
            aggregateBlocks.AddRange(titleDoc.Blocks);
            blocks.Add(new LatexProjectedBlock(titleCommand!.Syntax.Span, "title", titleDoc, titleDiagnostics));
        }
        BlockCandidate[] candidates = BuildCandidates(context)
            .OrderBy(static candidate => candidate.Span.Start.Offset)
            .ThenByDescending(static candidate => candidate.Span.Length).ToArray();
        int consumedUntil = document.Body?.ContentSpan.Start.Offset ?? 0;
        foreach (BlockCandidate candidate in candidates) {
            cancellationToken.ThrowIfCancellationRequested();
            if (candidate.Span.Start.Offset < consumedUntil) continue;
            var projected = MarkdownDoc.Create();
            var blockDiagnostics = new List<LatexMarkdownConversionDiagnostic>();
            AddCandidate(context, projected, candidate, options, blockDiagnostics);
            aggregateBlocks.AddRange(projected.Blocks);
            diagnostics.AddRange(blockDiagnostics);
            blocks.Add(new LatexProjectedBlock(candidate.Span, candidate.Kind, projected, blockDiagnostics));
            consumedUntil = candidate.Span.End.Offset;
        }
        context.CheckCancellation();
        target.AddRange(aggregateBlocks);
        context.CheckCancellation();
        return new LatexMarkdownProjection(document, new LatexToMarkdownResult(target, diagnostics), blocks, globalDiagnostics);
    }

    private static IEnumerable<BlockCandidate> BuildCandidates(LatexProjectionContext context) {
        if (context.Document.Body == null || context.Document.Profile == LatexDocumentProfile.PreserveOnly) {
            LatexSourceSpan fallback = context.Document.Body?.ContentSpan ?? context.Document.SyntaxTree.Root.Span;
            context.CheckCancellation();
            yield return new BlockCandidate(fallback, fallback);
            yield break;
        }
        int start = context.Document.Body.ContentSpan.Start.Offset;
        int end = context.Document.Body.ContentSpan.End.Offset;
        foreach (LatexHeading heading in context.Document.Headings.Where(heading => IsInside(heading.Command.Syntax.Span, start, end))) {
            context.CheckCancellation();
            yield return new BlockCandidate(heading.Command.Syntax.Span, heading);
        }
        foreach (LatexParagraph paragraph in context.Document.Paragraphs.Where(paragraph => IsInside(paragraph.Span, start, end))) {
            context.CheckCancellation();
            yield return new BlockCandidate(paragraph.Span, paragraph);
        }
        foreach (LatexList list in context.Document.Lists.Where(list => IsInside(list.Environment.Syntax.Span, start, end))) {
            context.CheckCancellation();
            yield return new BlockCandidate(list.Environment.Syntax.Span, list);
        }
        foreach (LatexFigure figure in context.Document.Figures.Where(figure => IsInside(figure.Environment.Syntax.Span, start, end))) {
            context.CheckCancellation();
            yield return new BlockCandidate(figure.Environment.Syntax.Span, figure);
        }
        foreach (LatexTable table in context.Document.Tables.Where(table => IsInside(table.Environment.Syntax.Span, start, end))) {
            LatexEnvironment? container = context.FindAncestorEnvironment(table.Environment, "table");
            context.CheckCancellation();
            yield return new BlockCandidate(container?.Syntax.Span ?? table.Environment.Syntax.Span, table);
        }
        foreach (LatexTheorem theorem in context.Document.Theorems.Where(theorem => IsInside(theorem.Environment.Syntax.Span, start, end))) {
            context.CheckCancellation();
            yield return new BlockCandidate(theorem.Environment.Syntax.Span, theorem);
        }
        foreach (LatexMath math in context.Document.Math.Where(math =>
                     math.Kind != LatexMathKind.InlineDollar && math.Kind != LatexMathKind.InlineParentheses &&
                     IsInside(math.Syntax.Span, start, end))) {
            context.CheckCancellation();
            yield return new BlockCandidate(math.Syntax.Span, math);
        }
        foreach (LatexSyntaxNode verbatim in context.Verbatim.Where(node =>
                     node.Kind == LatexSyntaxKind.Verbatim && LatexSemanticBuilder.IsActiveSyntax(node) && !LatexSemanticBuilder.IsInsideCommandArgument(node) &&
                     !string.Equals(node.Value, "verb", StringComparison.Ordinal) &&
                     IsInside(node.Span, start, end) && IsDirectChildSyntax(node, context.Document.Body.Syntax))) {
            context.CheckCancellation();
            yield return new BlockCandidate(verbatim.Span, verbatim);
        }
        foreach (LatexEnvironment environment in context.Document.Environments.Where(environment =>
                     !ReferenceEquals(environment, context.Document.Body) && IsInside(environment.Syntax.Span, start, end) &&
                     IsDirectChildEnvironment(environment.Syntax, context.Document.Body.Syntax) &&
                     !context.HasSemanticProjection(environment))) {
            context.CheckCancellation();
            yield return new BlockCandidate(environment.Syntax.Span, environment);
        }
    }

    private static void AddCandidate(
        LatexProjectionContext context,
        MarkdownDoc target,
        BlockCandidate candidate,
        LatexToMarkdownOptions options,
        List<LatexMarkdownConversionDiagnostic> diagnostics) {
        context.CheckCancellation();
        switch (candidate.Value) {
            case LatexSourceSpan fallback:
                AddSourceFallback(context, target, fallback, options, diagnostics);
                break;
            case LatexHeading heading: {
                    LatexArgument title = heading.Command.GetRequiredArgument(0)!;
                    int markdownLevel = GetMarkdownHeadingLevel(context, heading);
                    var block = new HeadingBlock(markdownLevel,
                        LatexInlineToMarkdownConverter.Convert(context, title.ContentSpan, diagnostics));
                    ApplyLabel(context, block, candidate.Span, diagnostics);
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
                    InlineSequence inlines = LatexInlineToMarkdownConverter.Convert(context, paragraph.Span, diagnostics);
                    if (inlines.Nodes.Count > 0) target.Add(new ParagraphBlock(inlines));
                    break;
                }
            case LatexList list:
                AddList(context, target, list, diagnostics);
                break;
            case LatexFigure figure:
                AddFigure(context, target, figure, options, diagnostics);
                break;
            case LatexTable table:
                target.Add(ConvertTable(context, table, diagnostics));
                LatexEnvironment? tableContainer = context.FindAncestorEnvironment(table.Environment, "table");
                if (tableContainer != null) {
                    var represented = new List<LatexSourceSpan> { table.Environment.Syntax.Span };
                    LatexCommand? caption = context.FindDirectCommand(tableContainer, "caption");
                    LatexCommand? label = context.FindDirectCommand(tableContainer, "label");
                    if (caption != null) represented.Add(caption.Syntax.Span);
                    if (label != null) represented.Add(label.Syntax.Span);
                    AddResidualSource(context, target, tableContainer, represented, options, diagnostics, "table-container");
                }
                break;
            case LatexTheorem theorem:
                AddTheorem(context, target, theorem, diagnostics);
                break;
            case LatexMath math:
                target.Add(new SemanticFencedBlock(MarkdownSemanticKinds.Math, "latex", ExtractVisibleSource(context, math.ContentSpan, diagnostics)));
                diagnostics.Add(new LatexMarkdownConversionDiagnostic(
                    "LATEXMD201", LatexMarkdownConversionOutcome.Simplified, "display-math",
                    "Display math source was transported without TeX layout evaluation.", math.Syntax.Span));
                break;
            case LatexEnvironment environment:
                AddEnvironmentFallback(context, target, environment, options, diagnostics);
                break;
            case LatexSyntaxNode verbatim:
                AddVerbatimBlock(context, target, verbatim, diagnostics);
                break;
        }
    }

    private static void AddVerbatimBlock(
        LatexProjectionContext context,
        MarkdownDoc target,
        LatexSyntaxNode syntax,
        List<LatexMarkdownConversionDiagnostic> diagnostics) {
        if (string.Equals(syntax.Value, "comment", StringComparison.Ordinal)) {
            diagnostics.Add(new LatexMarkdownConversionDiagnostic(
                "LATEXMD210", LatexMarkdownConversionOutcome.Omitted, "comment-environment",
                "The LaTeX comment environment was omitted and its body was not exposed as Markdown text.", syntax.Span));
            return;
        }
        target.Add(new CodeBlock("text", LatexInlineToMarkdownConverter.GetVerbatimContent(context, syntax)));
        diagnostics.Add(new LatexMarkdownConversionDiagnostic(
            "LATEXMD213", LatexMarkdownConversionOutcome.Simplified, "verbatim-environment",
            "Opaque LaTeX verbatim content was retained as a fenced code block without TeX environment semantics.", syntax.Span));
    }

    private static int GetMarkdownHeadingLevel(LatexProjectionContext context, LatexHeading heading) {
        int firstSectionLevel = string.Equals(context.Document.DocumentClassName, "article", StringComparison.Ordinal) ? 2 : 1;
        bool hasPart = context.HasPart;
        if (hasPart && heading.Level == 0) return 1;
        int markdownLevel = heading.Level - firstSectionLevel + 1;
        if (hasPart) markdownLevel++;
        return Math.Max(1, Math.Min(6, markdownLevel));
    }

    private static void AddList(
        LatexProjectionContext context,
        MarkdownDoc target,
        LatexList source,
        List<LatexMarkdownConversionDiagnostic> diagnostics) {
        if (source.Kind == LatexListKind.Description) {
            var definitions = new DefinitionListBlock();
            foreach (LatexListItem item in source.Items) {
                context.CheckCancellation();
                var term = item.ItemCommand.GetOptionalArgument(0) is LatexArgument label
                    ? LatexInlineToMarkdownConverter.Convert(context, label.ContentSpan, diagnostics)
                    : new InlineSequence { AutoSpacing = false };
                var body = new ParagraphBlock(LatexInlineToMarkdownConverter.Convert(context, item.ContentSpan, diagnostics));
                definitions.AddEntry(new DefinitionListEntry(term, new IMarkdownBlock[] { body }));
            }
            target.Add(definitions);
            return;
        }
        if (source.Kind == LatexListKind.Ordered) {
            var list = new OrderedListBlock();
            foreach (LatexListItem item in source.Items) {
                context.CheckCancellation();
                list.Items.Add(new ListItem(ConvertListItem(context, item, diagnostics)));
            }
            target.Add(list);
        } else {
            var list = new UnorderedListBlock();
            foreach (LatexListItem item in source.Items) {
                context.CheckCancellation();
                list.Items.Add(new ListItem(ConvertListItem(context, item, diagnostics)));
            }
            target.Add(list);
        }
    }

    private static InlineSequence ConvertListItem(
        LatexProjectionContext context,
        LatexListItem item,
        List<LatexMarkdownConversionDiagnostic> diagnostics) {
        InlineSequence content = LatexInlineToMarkdownConverter.Convert(context, item.ContentSpan, diagnostics);
        if (item.ItemCommand.GetOptionalArgument(0) is LatexArgument label) {
            InlineSequence labeled = LatexInlineToMarkdownConverter.Convert(context, label.ContentSpan, diagnostics);
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
        LatexProjectionContext context,
        MarkdownDoc target,
        LatexFigure source,
        LatexToMarkdownOptions options,
        List<LatexMarkdownConversionDiagnostic> diagnostics) {
        string caption = LatexInlineToMarkdownConverter.ReadDisplayArgument(context, source.CaptionCommand?.GetRequiredArgument(0), diagnostics, "figure-caption");
        string label = LatexInlineToMarkdownConverter.ReadArgumentSource(context, source.LabelCommand?.GetRequiredArgument(0), diagnostics);
        var images = new List<IMarkdownBlock>();
        for (int index = 0; index < source.Images.Count; index++) {
            context.CheckCancellation();
            LatexImage image = source.Images[index];
            LatexInlineToMarkdownConverter.ReportGraphicsOptions(image.Command, diagnostics);
            var block = new ImageBlock(LatexLiteralText.Decode(LatexInlineToMarkdownConverter.ReadArgumentSource(context, image.Command.GetRequiredArgument(0), diagnostics), context.CancellationToken), caption);
            if (!string.IsNullOrWhiteSpace(label)) block.SetAttributes(MarkdownAttributeSet.Create(label));
            if (!string.IsNullOrWhiteSpace(caption)) block.Caption = caption;
            images.Add(block);
        }
        target.AddRange(images);
        if (source.Images.Count == 0) {
            AddEnvironmentFallback(context, target, source.Environment, options, diagnostics);
            return;
        }
        var represented = source.Images.Select(static image => image.Command.Syntax.Span).ToList();
        if (source.CaptionCommand != null) represented.Add(source.CaptionCommand.Syntax.Span);
        if (source.LabelCommand != null) represented.Add(source.LabelCommand.Syntax.Span);
        AddResidualSource(context, target, source.Environment, represented, options, diagnostics, "figure-content");
    }

    private static TableBlock ConvertTable(
        LatexProjectionContext context,
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
            context.CheckCancellation();
            var values = new List<string>();
            var cells = new List<TableCell>();
            foreach (LatexTableCell sourceCell in source.Rows[rowIndex].Cells) {
                context.CheckCancellation();
                InlineSequence inlines = LatexInlineToMarkdownConverter.Convert(context, sourceCell.Span, diagnostics);
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
        LatexEnvironment? container = context.FindAncestorEnvironment(source.Environment, "table");
        LatexCommand? caption = context.FindDirectCommand(container, "caption");
        LatexCommand? label = context.FindDirectCommand(container, "label");
        string captionText = LatexInlineToMarkdownConverter.ReadDisplayArgument(context, caption?.GetRequiredArgument(0), diagnostics, "table-caption");
        string labelText = LatexInlineToMarkdownConverter.ReadArgumentSource(context, label?.GetRequiredArgument(0), diagnostics);
        if (!string.IsNullOrWhiteSpace(captionText) || !string.IsNullOrWhiteSpace(labelText)) {
            var attributes = string.IsNullOrWhiteSpace(captionText)
                ? null
                : new[] { new KeyValuePair<string, string?>("caption", captionText) };
            target.SetAttributes(MarkdownAttributeSet.Create(labelText, attributes: attributes));
        }
        return target;
    }

    private static void AddTheorem(
        LatexProjectionContext context,
        MarkdownDoc target,
        LatexTheorem source,
        List<LatexMarkdownConversionDiagnostic> diagnostics) {
        InlineSequence body = source.LabelCommand == null
            ? LatexInlineToMarkdownConverter.Convert(context, source.Environment.ContentSpan, diagnostics)
            : LatexInlineToMarkdownConverter.ConvertExcluding(
                context,
                source.Environment.ContentSpan,
                new[] { source.LabelCommand.Syntax.Span },
                diagnostics);
        var callout = new CalloutBlock(source.Kind, LatexInlineToMarkdownConverter.ReadDisplayArgument(context, source.Environment.BeginCommand.GetOptionalArgument(0), diagnostics, "theorem-title"),
            new IMarkdownBlock[] { new ParagraphBlock(body) });
        string label = LatexInlineToMarkdownConverter.ReadArgumentSource(context, source.LabelCommand?.GetRequiredArgument(0), diagnostics);
        if (!string.IsNullOrWhiteSpace(label)) callout.SetAttributes(MarkdownAttributeSet.Create(label));
        target.Add(callout);
    }

    private static void AddEnvironmentFallback(
        LatexProjectionContext context,
        MarkdownDoc target,
        LatexEnvironment source,
        LatexToMarkdownOptions options,
        List<LatexMarkdownConversionDiagnostic> diagnostics) {
        if (string.Equals(source.Name, "quote", StringComparison.Ordinal) || string.Equals(source.Name, "quotation", StringComparison.Ordinal)) {
            var quote = new QuoteBlock();
            quote.ChildBlocks.Add(new ParagraphBlock(LatexInlineToMarkdownConverter.Convert(context, source.ContentSpan, diagnostics)));
            target.Add(quote);
            return;
        }
        if (string.Equals(source.Name, "verbatim", StringComparison.Ordinal)) {
            target.Code("text", source.Content.Trim('\r', '\n'));
            return;
        }
        LatexSyntaxNode[] comments = context.Comments(source.Syntax.Span).Where(static item => item.OpaqueSyntax != null).Select(static item => item.OpaqueSyntax!).ToArray();
        if (options.PreserveUnsupportedAsSource) {
            string visibleSource = context.ExtractResidual(source.Syntax.Span, context.Comments(source.Syntax.Span).Select(static item => item.Span));
            if (!string.IsNullOrWhiteSpace(visibleSource)) target.Code("latex", visibleSource);
        }
        ReportOmittedComments(context, comments, diagnostics);
        diagnostics.Add(new LatexMarkdownConversionDiagnostic(
            "LATEXMD299",
            options.PreserveUnsupportedAsSource ? LatexMarkdownConversionOutcome.SourceFallback : LatexMarkdownConversionOutcome.Omitted,
            "environment:" + source.Name,
            options.PreserveUnsupportedAsSource ? "Unknown environment retained as visible LaTeX source." : "Unknown environment omitted by conversion options.",
            source.Syntax.Span));
    }

    private static void AddResidualSource(
        LatexProjectionContext context,
        MarkdownDoc target,
        LatexEnvironment environment,
        IEnumerable<LatexSourceSpan> representedSpans,
        LatexToMarkdownOptions options,
        List<LatexMarkdownConversionDiagnostic> diagnostics,
        string feature) {
        LatexSyntaxNode[] comments = context.Comments(environment.Syntax.Span).Where(static item => item.OpaqueSyntax != null).Select(static item => item.OpaqueSyntax!).ToArray();
        string residual = context.ExtractResidual(environment.ContentSpan, representedSpans.Concat(context.Comments(environment.ContentSpan).Select(static item => item.Span)));
        ReportOmittedComments(context, comments, diagnostics);
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

    private static void ReportOmittedComments(
        LatexProjectionContext context,
        IEnumerable<LatexSyntaxNode> comments,
        List<LatexMarkdownConversionDiagnostic> diagnostics) {
        foreach (LatexSyntaxNode comment in comments) {
            context.CheckCancellation();
            diagnostics.Add(new LatexMarkdownConversionDiagnostic(
                "LATEXMD210", LatexMarkdownConversionOutcome.Omitted, "comment-environment",
                "The LaTeX comment environment was omitted and its body was not exposed as Markdown text.",
                comment.Span));
        }
    }

    internal static string ExtractVisibleSource(LatexProjectionContext context, LatexSourceSpan span,
        List<LatexMarkdownConversionDiagnostic> diagnostics) {
        LatexProjectionComment[] comments = context.Comments(span).ToArray();
        ReportOmittedComments(context, comments.Where(static item => item.OpaqueSyntax != null).Select(static item => item.OpaqueSyntax!), diagnostics);
        return context.ExtractResidual(span, comments.Select(static item => item.Span));
    }

    private static void AddFrontMatter(LatexProjectionContext source, MarkdownDoc target, LatexToMarkdownOptions options,
        List<LatexMarkdownConversionDiagnostic> diagnostics) {
        if (!options.IncludePreambleAsFrontMatter) return;
        var values = new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase);
        if (source.Document.DocumentClassName != null) values["documentclass"] = source.Document.DocumentClassName;
        AddCommandValue(source, values, "title", diagnostics);
        AddCommandValue(source, values, "author", diagnostics);
        AddCommandValue(source, values, "date", diagnostics);
        if (values.Count > 0) target.FrontMatter(values);
    }

    private static void AddCommandValue(LatexProjectionContext source, Dictionary<string, object?> values, string name,
        List<LatexMarkdownConversionDiagnostic> diagnostics) {
        LatexArgument? argument = source.FirstCommand(name)?.GetRequiredArgument(0);
        string value = LatexInlineToMarkdownConverter.ReadDisplayArgument(source, argument, diagnostics, "metadata:" + name);
        if (!string.IsNullOrEmpty(value)) values[name] = value;
    }

    private static void ApplyLabel(LatexProjectionContext context, MarkdownObject target, LatexSourceSpan owner,
        List<LatexMarkdownConversionDiagnostic> diagnostics) {
        LatexLabel? label = context.FindAdjacentLabel(owner);
        if (label != null) target.SetAttributes(MarkdownAttributeSet.Create(LatexInlineToMarkdownConverter.ReadArgumentSource(context, label.Command.GetRequiredArgument(0), diagnostics)));
    }

    private static void AddSourceFallback(LatexProjectionContext context, MarkdownDoc target, LatexSourceSpan span,
        LatexToMarkdownOptions options, List<LatexMarkdownConversionDiagnostic> diagnostics) {
        LatexSyntaxNode[] comments = context.Comments(span).Where(static item => item.OpaqueSyntax != null).Select(static item => item.OpaqueSyntax!).ToArray();
        string visible = context.ExtractResidual(span, context.Comments(span).Select(static item => item.Span));
        ReportOmittedComments(context, comments.Where(item => IsInside(item.Span, span.Start.Offset, span.End.Offset)), diagnostics);
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
