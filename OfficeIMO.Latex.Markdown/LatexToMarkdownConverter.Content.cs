namespace OfficeIMO.Latex.Markdown;

internal static partial class LatexToMarkdownConverter {
    // One source-sorted structural index serves every list and quotation. Child
    // projection queries its own range and retains the original document offsets.
    internal static BlockCandidate[] IndexContentCandidates(LatexProjectionContext context) {
        if (context.Document.Body == null || context.Document.Profile == LatexDocumentProfile.PreserveOnly)
            return Array.Empty<BlockCandidate>();
        var candidates = BuildCandidates(context, includeInlineBodies: true).Where(candidate => candidate.Value is not LatexParagraph).ToList();
        var seen = new HashSet<LatexSyntaxNode>(candidates.Where(candidate => candidate.Value is LatexEnvironment)
            .Select(candidate => ((LatexEnvironment)candidate.Value).Syntax));
        foreach (LatexEnvironment environment in context.Document.Environments) {
            context.CheckCancellation();
            if (ReferenceEquals(environment, context.Document.Body) || context.HasSemanticProjection(environment) ||
                !LatexSemanticBuilder.IsActiveSyntax(environment.Syntax) || !seen.Add(environment.Syntax)) continue;
            candidates.Add(new BlockCandidate(environment.Syntax.Span, environment));
        }
        var verbatimSeen = new HashSet<LatexSyntaxNode>(candidates.Where(candidate => candidate.Value is LatexSyntaxNode)
            .Select(candidate => (LatexSyntaxNode)candidate.Value));
        foreach (LatexSyntaxNode verbatim in context.Verbatim) {
            context.CheckCancellation();
            if (verbatim.Value != "verb" && verbatimSeen.Add(verbatim)) candidates.Add(new BlockCandidate(verbatim.Span, verbatim));
        }
        return candidates.OrderBy(candidate => candidate.Span.Start.Offset)
            .ThenByDescending(candidate => candidate.Span.Length).ToArray();
    }

    internal static IReadOnlyList<IMarkdownBlock> ConvertContent(LatexProjectionContext context, LatexSourceSpan span,
        LatexToMarkdownOptions options, List<LatexMarkdownConversionDiagnostic> diagnostics) {
        var candidates = new List<BlockCandidate>();
        int cursor = span.Start.Offset;
        foreach (BlockCandidate candidate in context.ContentCandidates(span)) {
            context.CheckCancellation();
            if (candidate.Span.Start.Offset < cursor) continue;
            AddContentParagraphs(context, cursor, candidate.Span.Start.Offset, candidates);
            candidates.Add(candidate);
            cursor = candidate.Span.End.Offset;
        }
        AddContentParagraphs(context, cursor, span.End.Offset, candidates);
        var target = MarkdownDoc.Create();
        foreach (BlockCandidate candidate in WithCommandFallbacks(context, candidates.ToArray(), span)) {
            context.CheckCancellation();
            AddCandidate(context, target, candidate, options, diagnostics);
        }
        return target.Blocks.ToArray();
    }

    private static void AddContentParagraphs(LatexProjectionContext context, int start, int end, List<BlockCandidate> candidates) {
        if (end <= start) return;
        LatexSourceSpan[] inlineBodies = context.InlineCandidates(start, end)
            .Where(candidate => candidate.Command?.Name == "footnote").Select(candidate => candidate.Span).ToArray();
        foreach (LatexParagraph paragraph in LatexSemanticBuilder.BuildParagraphsInSpan(context.Document.Source,
                     context.Document.Source.CreateSpan(start, end), context.CancellationToken, inlineBodies))
            candidates.Add(new BlockCandidate(paragraph.Span, paragraph));
    }

    private static ListItem ConvertStructuredListItem(LatexProjectionContext context, LatexListItem item,
        LatexToMarkdownOptions options, List<LatexMarkdownConversionDiagnostic> diagnostics) {
        IReadOnlyList<IMarkdownBlock> blocks = ConvertContent(context, item.ContentSpan, options, diagnostics);
        InlineSequence content = blocks.FirstOrDefault() is ParagraphBlock paragraph
            ? paragraph.Inlines : new InlineSequence { AutoSpacing = false };
        int firstChild = blocks.FirstOrDefault() is ParagraphBlock ? 1 : 0;
        if (item.ItemCommand.GetOptionalArgument(0) is LatexArgument label) {
            InlineSequence labeled = LatexInlineToMarkdownConverter.Convert(context, label.ContentSpan, diagnostics);
            if (LatexProjectedText.HasVisibleText(labeled, context.CancellationToken) &&
                LatexProjectedText.HasVisibleText(content, context.CancellationToken)) labeled.AddRaw(new MarkdownTextRun(": "));
            foreach (IMarkdownInline inline in content.Nodes) labeled.AddRaw(inline);
            content = labeled;
            diagnostics.Add(new LatexMarkdownConversionDiagnostic("LATEXMD214", LatexMarkdownConversionOutcome.Simplified,
                "list-item-label", "The custom TeX item label was retained as visible item text; target list markers use Markdown numbering or bullets.",
                item.ItemCommand.Syntax.Span));
        }
        var result = new ListItem(content);
        for (int index = firstChild; index < blocks.Count; index++) result.NestedBlocks.Add(blocks[index]);
        result.ForceLoose = blocks.Count(block => block is ParagraphBlock) > 1;
        return result;
    }
}
