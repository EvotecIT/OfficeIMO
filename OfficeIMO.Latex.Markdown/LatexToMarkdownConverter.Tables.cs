namespace OfficeIMO.Latex.Markdown;

internal static partial class LatexToMarkdownConverter {
    private static TableBlock ConvertTable(
        LatexProjectionContext context,
        LatexTable source,
        List<LatexMarkdownConversionDiagnostic> diagnostics) {
        var target = new TableBlock();
        ApplyColumnAlignments(context, source, target, diagnostics);
        bool header = source.Rows.Count > 1 && source.Rows[0].Cells.All(static cell => cell.Content.TrimStart().StartsWith("\\textbf", StringComparison.Ordinal));
        // Markdown's object-tree binder pads every row to the effective column count.
        // Check the expanded shape before creating any projected cells.
        const int maximumProjectedColumns = 4096;
        int columns = Math.Max(target.Alignments.Count,
            source.Rows.Count == 0 ? 0 : source.Rows.Max(static row => row.Cells.Count));
        // Even an empty table materializes an aligned header when Reader asks for HeaderCells.
        long projectedRows = source.Rows.Count == 0 ? (columns > 0 ? 1L : 0L)
            : source.Rows.Count + (header ? 0L : 1L);
        if (columns > maximumProjectedColumns)
            throw new System.IO.InvalidDataException("LaTeX table exceeds the Markdown projection cell limit.");
        context.ChargeProjectedTableCells((long)columns * projectedRows);
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

    private static void ApplyColumnAlignments(LatexProjectionContext context, LatexTable source, TableBlock target,
        List<LatexMarkdownConversionDiagnostic> diagnostics) {
        string specification = source.ColumnSpecification;
        bool rules = false, unsupported = false;
        var alignments = new List<ColumnAlignment>();
        for (int index = 0; index < specification.Length; index++) {
            if ((index & 1023) == 0) context.CheckCancellation();
            char column = specification[index];
            if (char.IsWhiteSpace(column)) continue;
            if (column == '|') { rules = true; continue; }
            if (column == 'l') alignments.Add(ColumnAlignment.Left);
            else if (column == 'c') alignments.Add(ColumnAlignment.Center);
            else if (column == 'r') alignments.Add(ColumnAlignment.Right);
            else { unsupported = true; break; }
        }
        if (!unsupported) target.Alignments.AddRange(alignments);
        if (unsupported || rules) diagnostics.Add(new LatexMarkdownConversionDiagnostic(
            "LATEXMD215", LatexMarkdownConversionOutcome.Simplified, "table-columns",
            unsupported ? "The TeX column specification uses unsupported widths, modifiers, or repetition; target column alignment uses its defaults."
                : "Column alignment was retained; TeX vertical rule styling is not represented by Markdown pipe tables.",
            source.Environment.BeginCommand.GetRequiredArgument(1)?.ContentSpan ?? source.Environment.BeginCommand.Syntax.Span));
    }

}
