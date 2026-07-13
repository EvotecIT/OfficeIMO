namespace OfficeIMO.Latex.Markdown;

/// <summary>Converts the bounded OfficeIMO LaTeX profile to typed Markdown.</summary>
internal static class LatexToMarkdownConverter {
    /// <summary>Converts recognized semantics and diagnoses source fallbacks.</summary>
    internal static LatexToMarkdownResult Convert(
        LatexDocument document,
        LatexToMarkdownOptions? options = null) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        options ??= new LatexToMarkdownOptions();
        var target = MarkdownDoc.Create();
        var diagnostics = new List<LatexMarkdownConversionDiagnostic>();
        AddFrontMatter(document, target, options);

        LatexCommand? titleCommand = document.Commands.FirstOrDefault(static command => string.Equals(command.Name, "title", StringComparison.Ordinal));
        LatexArgument? title = titleCommand?.GetRequiredArgument(0);
        if (title != null) target.Add(new HeadingBlock(1, LatexInlineToMarkdownConverter.Convert(document, title.ContentSpan, diagnostics)));

        BlockCandidate[] candidates = BuildCandidates(document)
            .OrderBy(static candidate => candidate.Span.Start.Offset)
            .ThenByDescending(static candidate => candidate.Span.Length)
            .ToArray();
        int consumedUntil = document.Body?.ContentSpan.Start.Offset ?? 0;
        for (int index = 0; index < candidates.Length; index++) {
            BlockCandidate candidate = candidates[index];
            if (candidate.Span.Start.Offset < consumedUntil) continue;
            AddCandidate(document, target, candidate, options, diagnostics);
            consumedUntil = candidate.Span.End.Offset;
        }
        return new LatexToMarkdownResult(target, diagnostics);
    }

    private static IEnumerable<BlockCandidate> BuildCandidates(LatexDocument document) {
        if (document.Body == null) yield break;
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
        foreach (LatexEnvironment environment in document.Environments.Where(environment =>
                     !ReferenceEquals(environment, document.Body) && IsInside(environment.Syntax.Span, start, end) &&
                     IsDirectChildEnvironment(environment.Syntax, document.Body.Syntax) &&
                     !IsHandledEnvironment(environment.Name))) {
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
            case LatexHeading heading: {
                LatexArgument title = heading.Command.GetRequiredArgument(0)!;
                var block = new HeadingBlock(Math.Max(1, Math.Min(6, heading.Level)),
                    LatexInlineToMarkdownConverter.Convert(document, title.ContentSpan, diagnostics));
                ApplyLabel(document, block, candidate.Span);
                target.Add(block);
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
                target.Add(new SemanticFencedBlock(MarkdownSemanticKinds.Math, "latex", math.Content));
                diagnostics.Add(new LatexMarkdownConversionDiagnostic(
                    "LATEXMD201", LatexMarkdownConversionOutcome.Simplified, "display-math",
                    "Display math source was transported without TeX layout evaluation.", math.Syntax.Span));
                break;
            case LatexEnvironment environment:
                AddEnvironmentFallback(target, environment, options, diagnostics);
                break;
        }
    }

    private static void AddList(
        LatexDocument document,
        MarkdownDoc target,
        LatexList source,
        List<LatexMarkdownConversionDiagnostic> diagnostics) {
        if (source.Kind == LatexListKind.Description) {
            var definitions = new DefinitionListBlock();
            foreach (LatexListItem item in source.Items) {
                var term = new InlineSequence { AutoSpacing = false }.Text(item.Label ?? string.Empty);
                var body = new ParagraphBlock(LatexInlineToMarkdownConverter.Convert(document, item.ContentSpan, diagnostics));
                definitions.AddEntry(new DefinitionListEntry(term, new IMarkdownBlock[] { body }));
            }
            target.Add(definitions);
            return;
        }
        if (source.Kind == LatexListKind.Ordered) {
            var list = new OrderedListBlock();
            foreach (LatexListItem item in source.Items) {
                list.Items.Add(new ListItem(LatexInlineToMarkdownConverter.Convert(document, item.ContentSpan, diagnostics)));
            }
            target.Add(list);
        } else {
            var list = new UnorderedListBlock();
            foreach (LatexListItem item in source.Items) {
                list.Items.Add(new ListItem(LatexInlineToMarkdownConverter.Convert(document, item.ContentSpan, diagnostics)));
            }
            target.Add(list);
        }
    }

    private static void AddFigure(
        LatexDocument document,
        MarkdownDoc target,
        LatexFigure source,
        LatexToMarkdownOptions options,
        List<LatexMarkdownConversionDiagnostic> diagnostics) {
        for (int index = 0; index < source.Images.Count; index++) {
            LatexImage image = source.Images[index];
            var block = new ImageBlock(image.Target, source.Caption);
            if (!string.IsNullOrWhiteSpace(source.Label)) block.SetAttributes(MarkdownAttributeSet.Create(source.Label));
            if (!string.IsNullOrWhiteSpace(source.Caption)) block.Caption = source.Caption;
            target.Add(block);
        }
        if (source.Images.Count == 0) {
            AddEnvironmentFallback(target, source.Environment, options, diagnostics);
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
        string? captionText = caption?.GetRequiredArgument(0)?.Content;
        string? labelText = label?.GetRequiredArgument(0)?.Content;
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
        var callout = new CalloutBlock(source.Kind, source.Title ?? string.Empty,
            new IMarkdownBlock[] { new ParagraphBlock(body) });
        if (!string.IsNullOrWhiteSpace(source.Label)) callout.SetAttributes(MarkdownAttributeSet.Create(source.Label));
        target.Add(callout);
    }

    private static void AddEnvironmentFallback(
        MarkdownDoc target,
        LatexEnvironment source,
        LatexToMarkdownOptions options,
        List<LatexMarkdownConversionDiagnostic> diagnostics) {
        if (string.Equals(source.Name, "quote", StringComparison.Ordinal) || string.Equals(source.Name, "quotation", StringComparison.Ordinal)) {
            target.Quote(source.Content.Trim());
            return;
        }
        if (string.Equals(source.Name, "verbatim", StringComparison.Ordinal)) {
            target.Code("text", source.Content.Trim('\r', '\n'));
            return;
        }
        if (options.PreserveUnsupportedAsSource) target.Code("latex", source.Syntax.OriginalText);
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
        string residual = ExtractResidual(document.Source.Text, environment.ContentSpan, representedSpans);
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

    private static void AddFrontMatter(LatexDocument source, MarkdownDoc target, LatexToMarkdownOptions options) {
        if (!options.IncludePreambleAsFrontMatter) return;
        var values = new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase);
        if (source.DocumentClassName != null) values["documentclass"] = source.DocumentClassName;
        AddCommandValue(source, values, "title");
        AddCommandValue(source, values, "author");
        AddCommandValue(source, values, "date");
        if (values.Count > 0) target.FrontMatter(values);
    }

    private static void AddCommandValue(LatexDocument source, Dictionary<string, object?> values, string name) {
        string? value = source.Commands.FirstOrDefault(command => string.Equals(command.Name, name, StringComparison.Ordinal))?.GetRequiredArgument(0)?.Content;
        if (!string.IsNullOrEmpty(value)) values[name] = value;
    }

    private static void ApplyLabel(LatexDocument document, MarkdownObject target, LatexSourceSpan owner) {
        LatexLabel? label = FindAdjacentLabel(document, owner);
        if (label != null) target.SetAttributes(MarkdownAttributeSet.Create(label.Name));
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

    private static bool IsHandledEnvironment(string name) =>
        name == "itemize" || name == "enumerate" || name == "description" || name == "figure" || name == "table" || name == "tabular" ||
        name == "equation" || name == "equation*" || name == "align" || name == "align*" || name == "gather" || name == "gather*" ||
        name == "multline" || name == "multline*" || name == "theorem" || name == "lemma" || name == "proposition" ||
        name == "corollary" || name == "definition" || name == "remark" || name == "proof";

    private static bool IsDirectChildEnvironment(LatexSyntaxNode node, LatexSyntaxNode body) {
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
    }
}

/// <summary>Conversion extensions.</summary>
public static partial class LatexMarkdownConverterExtensions {
    /// <summary>Converts a native LaTeX document to Markdown.</summary>
    public static LatexToMarkdownResult ToMarkdownDocumentResult(this LatexDocument document, LatexToMarkdownOptions? options = null) =>
        LatexToMarkdownConverter.Convert(document, options);

    /// <summary>Converts a LaTeX document to a typed Markdown document.</summary>
    public static MarkdownDoc ToMarkdownDocument(this LatexDocument document, LatexToMarkdownOptions? options = null) =>
        document.ToMarkdownDocumentResult(options).Value;
}
