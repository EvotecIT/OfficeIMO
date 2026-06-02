using AngleSharp.Dom;
using OfficeIMO.Markdown;

namespace OfficeIMO.Markdown.Html;

public sealed partial class HtmlToMarkdownConverter {
    private static TableBlock ConvertTableElement(IElement element, ConversionContext context) {
        var table = new TableBlock();
        bool headerWritten = false;
        var headerCells = new List<TableCell>();
        var rowCells = new List<IReadOnlyList<TableCell>>();

        foreach (var row in EnumerateTableRows(element)) {
            var cells = row.Children
                .Where(child => child.TagName.Equals("TH", StringComparison.OrdinalIgnoreCase) || child.TagName.Equals("TD", StringComparison.OrdinalIgnoreCase))
                .ToList();
            if (cells.Count == 0) {
                continue;
            }

            bool isHeaderRow = !headerWritten && cells.All(cell => cell.TagName.Equals("TH", StringComparison.OrdinalIgnoreCase));
            var renderedCells = new List<string>(cells.Count);
            var structuredCells = new List<TableCell>(cells.Count);
            foreach (var cell in cells) {
                var cellBlocks = ConvertTableCellToBlocks(cell, context);
                structuredCells.Add(new TableCell(cellBlocks));
                renderedCells.Add(RenderTableCellBlocksToMarkdown(cellBlocks));
                if (isHeaderRow) {
                    table.Alignments.Add(ParseAlignment(cell));
                }
            }

            if (isHeaderRow) {
                foreach (var value in renderedCells) {
                    table.Headers.Add(value);
                }
                headerCells.AddRange(structuredCells);
                headerWritten = true;
            } else {
                table.Rows.Add(renderedCells);
                rowCells.Add(structuredCells);
            }
        }

        if (!headerWritten && table.Rows.Count > 0) {
            var firstRow = table.Rows[0];
            table.Rows.RemoveAt(0);
            var firstStructuredRow = rowCells[0];
            rowCells.RemoveAt(0);
            foreach (var value in firstRow) {
                table.Headers.Add(value);
                table.Alignments.Add(ColumnAlignment.None);
            }
            headerCells.AddRange(firstStructuredRow);
        }

        table.SetStructuredCells(headerCells, rowCells, table.ComputeContentSignature());

        return table;
    }

    private static IEnumerable<IElement> EnumerateTableRows(IElement table) {
        foreach (var child in table.Children) {
            if (child.TagName.Equals("TR", StringComparison.OrdinalIgnoreCase)) {
                yield return child;
                continue;
            }

            if (!child.TagName.Equals("THEAD", StringComparison.OrdinalIgnoreCase)
                && !child.TagName.Equals("TBODY", StringComparison.OrdinalIgnoreCase)
                && !child.TagName.Equals("TFOOT", StringComparison.OrdinalIgnoreCase)) {
                continue;
            }

            foreach (var row in child.Children.Where(static row => row.TagName.Equals("TR", StringComparison.OrdinalIgnoreCase))) {
                yield return row;
            }
        }
    }

    private static ColumnAlignment ParseAlignment(IElement cell) {
        var alignment = ParseAlignmentValue(cell.GetAttribute("align"));
        if (alignment != ColumnAlignment.None) {
            return alignment;
        }

        return ParseAlignmentValue(TryGetStyleDeclarationValue(cell.GetAttribute("style"), "text-align"));
    }

    private static ColumnAlignment ParseAlignmentValue(string? rawAlignment) {
        if (string.IsNullOrWhiteSpace(rawAlignment)) {
            return ColumnAlignment.None;
        }

        switch (rawAlignment!.Trim().ToLowerInvariant()) {
            case "left":
                return ColumnAlignment.Left;
            case "center":
                return ColumnAlignment.Center;
            case "right":
                return ColumnAlignment.Right;
            default:
                return ColumnAlignment.None;
        }
    }

    private static IReadOnlyList<IMarkdownBlock> ConvertTableCellToBlocks(IElement cell, ConversionContext context) {
        if (HasDirectBlockChildren(cell, context)) {
            return ConvertNodesToBlocks(cell.ChildNodes, context);
        }

        var inlineSequence = NormalizeInlineSequenceForBlock(ConvertInlineNodesToInlineSequence(cell.ChildNodes, context));
        if (!HasVisibleInlineContent(inlineSequence)) {
            return Array.Empty<IMarkdownBlock>();
        }

        return new IMarkdownBlock[] { new ParagraphBlock(inlineSequence) };
    }

    private static string RenderTableCellBlocksToMarkdown(IReadOnlyList<IMarkdownBlock> blocks) {
        if (blocks == null || blocks.Count == 0) {
            return string.Empty;
        }

        return new TableCell(blocks).Markdown.Replace("  \n", "<br>");
    }

}
