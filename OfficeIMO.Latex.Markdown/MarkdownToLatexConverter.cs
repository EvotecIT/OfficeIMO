namespace OfficeIMO.Latex.Markdown;

/// <summary>Generates canonical bounded-profile LaTeX from typed Markdown.</summary>
internal static class MarkdownToLatexConverter {
    /// <summary>Converts a Markdown semantic document and reparses generated LaTeX losslessly.</summary>
    internal static MarkdownToLatexResult Convert(
        MarkdownDoc document,
        MarkdownToLatexOptions? options = null) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        options ??= new MarkdownToLatexOptions();
        Validate(options);
        var diagnostics = new List<LatexMarkdownConversionDiagnostic>();
        var state = new ConversionState(options, document.Blocks.OfType<FootnoteDefinitionBlock>());

        string? title = GetFrontMatter(document, "title");
        HeadingBlock? firstHeading = options.FirstHeadingIsTitle
            ? document.Blocks.OfType<HeadingBlock>().FirstOrDefault(static heading => heading.Level == 1)
            : null;
        HeadingBlock? titleHeading = null;
        if (title == null && firstHeading != null) {
            title = firstHeading.Text;
            titleHeading = firstHeading;
        } else if (title != null && firstHeading != null && string.Equals(title, firstHeading.Text, StringComparison.Ordinal)) {
            titleHeading = firstHeading;
        }
        string? author = GetFrontMatter(document, "author");
        string? date = GetFrontMatter(document, "date");

        var body = new List<string>();
        for (int index = 0; index < document.Blocks.Count; index++) {
            IMarkdownBlock block = document.Blocks[index];
            if (title != null && !state.TitleConsumed && ReferenceEquals(block, titleHeading)) {
                state.TitleConsumed = true;
                body.Add("\\maketitle");
                continue;
            }
            string converted = ConvertBlock(block, options, state, diagnostics);
            if (converted.Length > 0) body.Add(converted);
        }
        if (title != null && !state.TitleConsumed) body.Insert(0, "\\maketitle");

        var source = new StringBuilder();
        source.Append("\\documentclass{").Append(options.DocumentClass).Append('}').Append(options.LineEnding);
        foreach (string package in state.Packages.OrderBy(static package => package, StringComparer.Ordinal)) {
            source.Append("\\usepackage{").Append(package).Append('}').Append(options.LineEnding);
        }
        foreach (string theorem in state.TheoremEnvironments.OrderBy(static value => value, StringComparer.Ordinal)) {
            source.Append("\\newtheorem{").Append(theorem).Append("}{").Append(TheoremDisplayName(theorem)).Append('}').Append(options.LineEnding);
        }
        if (title != null) source.Append("\\title{").Append(MarkdownInlineToLatexConverter.EscapeText(title)).Append('}').Append(options.LineEnding);
        if (author != null) source.Append("\\author{").Append(MarkdownInlineToLatexConverter.EscapeText(author)).Append('}').Append(options.LineEnding);
        if (date != null) source.Append("\\date{").Append(MarkdownInlineToLatexConverter.EscapeText(date)).Append('}').Append(options.LineEnding);
        source.Append("\\begin{document}").Append(options.LineEnding);
        if (body.Count > 0) source.Append(options.LineEnding).Append(string.Join(options.LineEnding + options.LineEnding, body)).Append(options.LineEnding);
        source.Append("\\end{document}").Append(options.LineEnding);

        string value = source.ToString();
        LatexDocument parsed = LatexDocument.Parse(value).Document;
        return new MarkdownToLatexResult(value, parsed, diagnostics);
    }

    private static string ConvertBlock(
        IMarkdownBlock block,
        MarkdownToLatexOptions options,
        ConversionState state,
        List<LatexMarkdownConversionDiagnostic> diagnostics) {
        switch (block) {
            case HeadingBlock heading:
                string command = HeadingCommand(heading.Level, options.FirstHeadingIsTitle);
                return "\\" + command + "{" + MarkdownInlineToLatexConverter.Convert(heading.Inlines, state, diagnostics, heading) + "}" + Label(heading, diagnostics);
            case ParagraphBlock paragraph:
                return MarkdownInlineToLatexConverter.Convert(paragraph.Inlines, state, diagnostics, paragraph) + Label(paragraph, diagnostics);
            case UnorderedListBlock unordered:
                return ConvertList("itemize", unordered.Items, options, state, diagnostics);
            case OrderedListBlock ordered:
                return ConvertList("enumerate", ordered.Items, options, state, diagnostics);
            case DefinitionListBlock definitions:
                return ConvertDefinitions(definitions, options, state, diagnostics);
            case CodeBlock code:
                return ConvertVerbatim(code.Content, code, diagnostics, options.LineEnding);
            case SemanticFencedBlock semantic when string.Equals(semantic.SemanticKind, MarkdownSemanticKinds.Math, StringComparison.OrdinalIgnoreCase):
                state.Packages.Add("amsmath");
                return "\\begin{equation*}" + options.LineEnding + semantic.Content + Ending(semantic.Content, options.LineEnding) + "\\end{equation*}";
            case TableBlock table:
                return ConvertTable(table, options, state, diagnostics);
            case ImageBlock image:
                return ConvertImage(image, options, state, diagnostics);
            case CalloutBlock callout:
                return ConvertCallout(callout, options, state, diagnostics);
            case QuoteBlock quote:
                return ConvertQuote(quote, options, state, diagnostics);
            case FootnoteDefinitionBlock:
                return string.Empty;
            default:
                return ConvertUnsupported(block, options, diagnostics);
        }
    }

    internal static string ConvertFootnoteReference(
        FootnoteRefInline source,
        MarkdownObject owner,
        ConversionState state,
        List<LatexMarkdownConversionDiagnostic> diagnostics) {
        if (!state.TryBeginFootnote(source.Label, out FootnoteDefinitionBlock? definition)) {
            diagnostics.Add(new LatexMarkdownConversionDiagnostic(
                "MDLATEX104", LatexMarkdownConversionOutcome.SourceFallback, "footnote-reference",
                "Markdown footnote reference has no resolvable definition and was retained as visible source.", null, owner.SourceSpan));
            return "\\texttt{" + MarkdownInlineToLatexConverter.EscapeText("[^" + source.Label + "]") + "}";
        }

        try {
            IReadOnlyList<IMarkdownBlock> blocks = definition!.ChildBlocks;
            if (blocks.Count == 0) return MarkdownInlineToLatexConverter.EscapeText(definition.Text);
            return string.Join(state.LineEnding + state.LineEnding,
                blocks.Select(block => ConvertBlock(block, state.Options, state, diagnostics)).Where(static value => value.Length > 0));
        } finally {
            state.EndFootnote(source.Label);
        }
    }

    private static string ConvertList(
        string environment,
        IReadOnlyList<ListItem> items,
        MarkdownToLatexOptions options,
        ConversionState state,
        List<LatexMarkdownConversionDiagnostic> diagnostics) {
        var output = new StringBuilder();
        output.Append("\\begin{").Append(environment).Append('}').Append(options.LineEnding);
        for (int index = 0; index < items.Count; index++) {
            ListItem item = items[index];
            output.Append("\\item ").Append(MarkdownInlineToLatexConverter.Convert(item.Content, state, diagnostics, item));
            for (int childIndex = 0; childIndex < item.NestedBlocks.Count; childIndex++) {
                output.Append(options.LineEnding).Append(ConvertBlock(item.NestedBlocks[childIndex], options, state, diagnostics));
            }
            output.Append(options.LineEnding);
        }
        output.Append("\\end{").Append(environment).Append('}');
        return output.ToString();
    }

    private static string ConvertDefinitions(
        DefinitionListBlock source,
        MarkdownToLatexOptions options,
        ConversionState state,
        List<LatexMarkdownConversionDiagnostic> diagnostics) {
        var output = new StringBuilder("\\begin{description}").Append(options.LineEnding);
        foreach (DefinitionListEntry entry in source.Entries) {
            output.Append("\\item[").Append(MarkdownInlineToLatexConverter.Convert(entry.Term, state, diagnostics, entry)).Append("] ");
            for (int index = 0; index < entry.DefinitionBlocks.Count; index++) {
                if (index > 0) output.Append(options.LineEnding).Append(options.LineEnding);
                output.Append(ConvertBlock(entry.DefinitionBlocks[index], options, state, diagnostics));
            }
            output.Append(options.LineEnding);
        }
        output.Append("\\end{description}");
        return output.ToString();
    }

    private static string ConvertTable(
        TableBlock source,
        MarkdownToLatexOptions options,
        ConversionState state,
        List<LatexMarkdownConversionDiagnostic> diagnostics) {
        TableRow[] rows = source.EnumerateRows().ToArray();
        int columns = rows.Length == 0
            ? 1
            : Math.Max(1, rows.Max(static row => row.Cells.Sum(static cell => Math.Max(1, cell.ColumnSpan))));
        string? caption = source.Attributes.GetAttribute("caption");
        bool useTableEnvironment = !string.IsNullOrWhiteSpace(caption) || !string.IsNullOrWhiteSpace(source.Attributes.ElementId);
        var output = new StringBuilder();
        if (useTableEnvironment) {
            output.Append("\\begin{table}").Append(options.LineEnding);
            if (!string.IsNullOrWhiteSpace(caption)) {
                output.Append("\\caption{").Append(MarkdownInlineToLatexConverter.EscapeText(caption!)).Append('}').Append(options.LineEnding);
            }
            output.Append(Label(source, diagnostics));
            if (!string.IsNullOrWhiteSpace(source.Attributes.ElementId)) output.Append(options.LineEnding);
        }
        output.Append("\\begin{tabular}{").Append(new string('l', columns)).Append('}').Append(options.LineEnding);
        for (int rowIndex = 0; rowIndex < rows.Length; rowIndex++) {
            for (int cellIndex = 0; cellIndex < rows[rowIndex].Cells.Count; cellIndex++) {
                if (cellIndex > 0) output.Append(" & ");
                TableCell cell = rows[rowIndex].Cells[cellIndex];
                string content = ConvertCell(cell, options, state, diagnostics);
                if (cell.RowSpan > 1) { state.Packages.Add("multirow"); content = "\\multirow{" + cell.RowSpan + "}{*}{" + content + "}"; }
                if (cell.ColumnSpan > 1) content = "\\multicolumn{" + cell.ColumnSpan + "}{l}{" + content + "}";
                output.Append(content);
            }
            output.Append(" \\\\").Append(options.LineEnding);
            if (rows[rowIndex].IsHeader) output.Append("\\hline").Append(options.LineEnding);
        }
        output.Append("\\end{tabular}");
        if (useTableEnvironment) output.Append(options.LineEnding).Append("\\end{table}");
        return output.ToString();
    }

    private static string ConvertCell(
        TableCell cell,
        MarkdownToLatexOptions options,
        ConversionState state,
        List<LatexMarkdownConversionDiagnostic> diagnostics) {
        if (cell.ChildBlocks.Count == 1 && cell.ChildBlocks[0] is ParagraphBlock paragraph) {
            string value = MarkdownInlineToLatexConverter.Convert(paragraph.Inlines, state, diagnostics, paragraph);
            if (cell.Bold || cell.IsHeader) value = "\\textbf{" + value + "}";
            if (cell.Italic) value = "\\emph{" + value + "}";
            return value;
        }
        return string.Join(" ", cell.ChildBlocks.Select(block => ConvertBlock(block, options, state, diagnostics)));
    }

    private static string ConvertImage(
        ImageBlock source,
        MarkdownToLatexOptions options,
        ConversionState state,
        List<LatexMarkdownConversionDiagnostic> diagnostics) {
        state.Packages.Add("graphicx");
        var output = new StringBuilder("\\begin{figure}").Append(options.LineEnding)
            .Append("\\includegraphics{").Append(MarkdownInlineToLatexConverter.EscapeArgument(source.Path)).Append('}').Append(options.LineEnding);
        if (!string.IsNullOrWhiteSpace(source.Caption)) output.Append("\\caption{").Append(MarkdownInlineToLatexConverter.EscapeText(source.Caption!)).Append('}').Append(options.LineEnding);
        if (!string.IsNullOrWhiteSpace(source.Attributes.ElementId)) output.Append(Label(source, diagnostics)).Append(options.LineEnding);
        output.Append("\\end{figure}");
        return output.ToString();
    }

    private static string ConvertCallout(
        CalloutBlock source,
        MarkdownToLatexOptions options,
        ConversionState state,
        List<LatexMarkdownConversionDiagnostic> diagnostics) {
        string theoremKind = source.Kind.ToLowerInvariant();
        if (IsTheorem(theoremKind)) {
            state.Packages.Add("amsthm");
            if (!string.Equals(theoremKind, "proof", StringComparison.Ordinal)) state.TheoremEnvironments.Add(theoremKind);
            var output = new StringBuilder("\\begin{").Append(theoremKind).Append('}');
            if (!string.IsNullOrWhiteSpace(source.Title)) output.Append('[').Append(MarkdownInlineToLatexConverter.EscapeText(source.Title)).Append(']');
            output.Append(options.LineEnding);
            if (!string.IsNullOrWhiteSpace(source.Attributes.ElementId)) output.Append(Label(source, diagnostics)).Append(options.LineEnding);
            foreach (IMarkdownBlock child in source.ChildBlocks) output.Append(ConvertBlock(child, options, state, diagnostics)).Append(options.LineEnding);
            output.Append("\\end{").Append(theoremKind).Append('}');
            return output.ToString();
        }
        var quote = new StringBuilder("\\begin{quote}").Append(options.LineEnding)
            .Append("\\textbf{").Append(MarkdownInlineToLatexConverter.EscapeText(source.Kind.ToUpperInvariant())).Append(":} ");
        foreach (IMarkdownBlock child in source.ChildBlocks) quote.Append(ConvertBlock(child, options, state, diagnostics)).Append(options.LineEnding);
        quote.Append("\\end{quote}");
        return quote.ToString();
    }

    private static string ConvertQuote(
        QuoteBlock source,
        MarkdownToLatexOptions options,
        ConversionState state,
        List<LatexMarkdownConversionDiagnostic> diagnostics) {
        var body = new List<string>();
        if (source.ChildBlocks.Count > 0) body.AddRange(source.ChildBlocks.Select(block => ConvertBlock(block, options, state, diagnostics)));
        else body.AddRange(source.Lines.Select(MarkdownInlineToLatexConverter.EscapeText));
        return "\\begin{quote}" + options.LineEnding + string.Join(options.LineEnding, body) + options.LineEnding + "\\end{quote}";
    }

    private static string ConvertVerbatim(
        string content,
        MarkdownObject owner,
        List<LatexMarkdownConversionDiagnostic> diagnostics,
        string lineEnding) {
        string value = content;
        if (value.IndexOf("\\end{verbatim}", StringComparison.Ordinal) >= 0) {
            value = value.Replace("\\end{verbatim}", "\\textbackslash{}end\\{verbatim\\}");
            diagnostics.Add(new LatexMarkdownConversionDiagnostic(
                "MDLATEX021", LatexMarkdownConversionOutcome.Simplified, "verbatim-delimiter",
                "A verbatim closing delimiter inside code was escaped.", null, owner.SourceSpan));
        }
        return "\\begin{verbatim}" + lineEnding + value + Ending(value, lineEnding) + "\\end{verbatim}";
    }

    private static string ConvertUnsupported(
        IMarkdownBlock source,
        MarkdownToLatexOptions options,
        List<LatexMarkdownConversionDiagnostic> diagnostics) {
        MarkdownObject? owner = source as MarkdownObject;
        diagnostics.Add(new LatexMarkdownConversionDiagnostic(
            "MDLATEX099",
            options.PreserveUnsupportedAsSource ? LatexMarkdownConversionOutcome.SourceFallback : LatexMarkdownConversionOutcome.Omitted,
            source.GetType().Name,
            options.PreserveUnsupportedAsSource ? "Unsupported Markdown retained in a verbatim environment." : "Unsupported Markdown omitted.",
            null,
            owner?.SourceSpan));
        return options.PreserveUnsupportedAsSource
            ? ConvertVerbatim(source.RenderMarkdown(), owner ?? new ParagraphBlock(new InlineSequence()), diagnostics, options.LineEnding)
            : string.Empty;
    }

    private static string HeadingCommand(int level, bool firstHeadingIsTitle) {
        int adjusted = firstHeadingIsTitle ? Math.Max(1, level - 1) : level;
        switch (adjusted) {
            case 1: return "section";
            case 2: return "subsection";
            case 3: return "subsubsection";
            case 4: return "paragraph";
            default: return "subparagraph";
        }
    }

    private static string Label(
        MarkdownObject source,
        List<LatexMarkdownConversionDiagnostic> diagnostics) {
        if (string.IsNullOrWhiteSpace(source.Attributes.ElementId)) return string.Empty;
        string original = source.Attributes.ElementId!;
        string normalized = MarkdownInlineToLatexConverter.EscapeLabel(original);
        if (!string.Equals(original, normalized, StringComparison.Ordinal)) {
            diagnostics.Add(new LatexMarkdownConversionDiagnostic(
                "MDLATEX031", LatexMarkdownConversionOutcome.Simplified, "label",
                "A Markdown identifier was encoded to a TeX-safe label.", null, source.SourceSpan));
        }
        return "\\label{" + normalized + "}";
    }

    private static string? GetFrontMatter(MarkdownDoc document, string key) => document.DocumentHeader?.FindEntry(key)?.Value?.ToString();

    private static bool IsTheorem(string kind) =>
        kind == "theorem" || kind == "lemma" || kind == "proposition" || kind == "corollary" ||
        kind == "definition" || kind == "remark" || kind == "proof";

    private static string TheoremDisplayName(string kind) => kind.Length == 0
        ? "Theorem"
        : char.ToUpperInvariant(kind[0]) + kind.Substring(1);

    private static string Ending(string value, string lineEnding) =>
        value.EndsWith("\n", StringComparison.Ordinal) || value.EndsWith("\r", StringComparison.Ordinal) || value.Length == 0 ? string.Empty : lineEnding;

    private static void Validate(MarkdownToLatexOptions options) {
        if (options.LineEnding != "\n" && options.LineEnding != "\r" && options.LineEnding != "\r\n") throw new ArgumentException("LineEnding must be LF, CR, or CRLF.");
        if (string.IsNullOrWhiteSpace(options.DocumentClass) || options.DocumentClass.Any(static character => !char.IsLetter(character))) {
            throw new ArgumentException("DocumentClass must be a simple alphabetic class name.");
        }
    }
}

public static partial class LatexMarkdownConverterExtensions {
    /// <summary>Converts Markdown to canonical bounded-profile LaTeX.</summary>
    public static MarkdownToLatexResult ToLatexDocumentResult(this MarkdownDoc document, MarkdownToLatexOptions? options = null) =>
        MarkdownToLatexConverter.Convert(document, options);

    /// <summary>Converts a Markdown document to a parsed canonical LaTeX document.</summary>
    public static LatexDocument ToLatexDocument(this MarkdownDoc document, MarkdownToLatexOptions? options = null) =>
        document.ToLatexDocumentResult(options).Value;
}
