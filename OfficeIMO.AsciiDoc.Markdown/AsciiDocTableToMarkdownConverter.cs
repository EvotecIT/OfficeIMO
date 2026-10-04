namespace OfficeIMO.AsciiDoc.Markdown;

internal static class AsciiDocTableToMarkdownConverter {
    internal static TableBlock Convert(
        AsciiDocTableBlock source,
        AsciiDocDocumentAttributes attributes,
        AsciiDocToMarkdownOptions options,
        List<AsciiDocMarkdownConversionDiagnostic> diagnostics,
        int depth) {
        var target = new TableBlock();
        var structuredHeaders = new List<TableCell>();
        var structuredRows = new List<IReadOnlyList<TableCell>>();
        bool hasHeader = source.Table.Rows.Count > 0 && source.Table.Rows[0].IsHeader;

        for (int rowIndex = 0; rowIndex < source.Table.Rows.Count; rowIndex++) {
            AsciiDocTableRow row = source.Table.Rows[rowIndex];
            var values = new List<string>();
            var cells = new List<TableCell>();
            for (int cellIndex = 0; cellIndex < row.Cells.Count; cellIndex++) {
                AsciiDocTableCell sourceCell = row.Cells[cellIndex];
                values.Add(sourceCell.Value);
                IReadOnlyList<IMarkdownBlock> blocks;
                if (sourceCell.GetBody() is AsciiDocDocument body) {
                    int remaining = options.MaximumBlockNestingDepth - depth - 1;
                    if (remaining < 1) throw new System.IO.InvalidDataException("AsciiDoc table conversion exceeds MaximumBlockNestingDepth.");
                    var content = MarkdownDoc.Create();
                    var attached = new HashSet<AsciiDocBlock>(body.BlocksOfType<AsciiDocListBlock>().SelectMany(list => list.Items).SelectMany(item => item.AttachedBlocks));
                    foreach (AsciiDocBlockContext context in body.GetBlockContextsFromSnapshot(attributes, true, remaining))
                        if (!attached.Contains(context.Block)) AsciiDocToMarkdownConverter.AddBlock(content, context.Block, context.Attributes, options, diagnostics, depth + 1);
                    blocks = content.Blocks.ToArray();
                } else if (sourceCell.Style == 'l') blocks = new IMarkdownBlock[] { new CodeBlock(string.Empty, sourceCell.Value) };
                else {
                    InlineSequence inlines;
                    if (sourceCell.Style == 'm') { inlines = new InlineSequence { AutoSpacing = false }; inlines.AddRaw(new CodeSpanInline(sourceCell.Value)); }
                    else inlines = AsciiDocInlineToMarkdownConverter.Convert(sourceCell.Inlines!, attributes, options, diagnostics, source);
                    blocks = new IMarkdownBlock[] { new ParagraphBlock(inlines) };
                }
                var targetCell = new TableCell(blocks) {
                    ColumnSpan = sourceCell.ColumnSpan,
                    RowSpan = sourceCell.RowSpan,
                    Bold = sourceCell.Style == 's' || sourceCell.Style == 'h',
                    Italic = sourceCell.Style == 'e'
                };
                cells.Add(targetCell);
            }
            while (values.Count < source.Table.ColumnCount) values.Add(string.Empty);
            if (hasHeader && rowIndex == 0) {
                target.Headers.AddRange(values);
                structuredHeaders.AddRange(cells);
            } else {
                target.Rows.Add(values);
                structuredRows.Add(cells);
            }
        }

        target.SetStructuredCells(structuredHeaders, structuredRows, target.ComputeContentSignature());
        if (source.Table.Cells.Any(static cell => cell.ColumnSpan > 1 || cell.RowSpan > 1)) {
            diagnostics.Add(new AsciiDocMarkdownConversionDiagnostic(
                "ADOCMD041",
                AsciiDocMarkdownDiagnosticSeverity.Warning,
                AsciiDocMarkdownConversionOutcome.Simplified,
                "table-spans",
                "Cell spans are retained in the typed Markdown table for rich targets; plain pipe Markdown cannot represent them.",
                source.Span));
        }
        return target;
    }

}
