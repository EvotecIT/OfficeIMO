namespace OfficeIMO.AsciiDoc.Markdown;

internal static class AsciiDocTableToMarkdownConverter {
    internal static TableBlock Convert(
        AsciiDocTableBlock source,
        AsciiDocDocumentAttributes attributes,
        AsciiDocToMarkdownOptions options,
        List<AsciiDocMarkdownConversionDiagnostic> diagnostics) {
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
                InlineSequence inlines = ParseCellInlines(sourceCell.Value, attributes, options, diagnostics, source);
                var targetCell = new TableCell(new IMarkdownBlock[] { new ParagraphBlock(inlines) }) {
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

    private static InlineSequence ParseCellInlines(
        string value,
        AsciiDocDocumentAttributes attributes,
        AsciiDocToMarkdownOptions options,
        List<AsciiDocMarkdownConversionDiagnostic> diagnostics,
        AsciiDocBlock owner) {
        AsciiDocParagraph? paragraph = AsciiDocDocument.Parse(value).Document.BlocksOfType<AsciiDocParagraph>().FirstOrDefault();
        return paragraph == null
            ? new InlineSequence { AutoSpacing = false }.Text(value)
            : AsciiDocInlineToMarkdownConverter.Convert(paragraph.Inlines, attributes, options, diagnostics, owner);
    }
}
