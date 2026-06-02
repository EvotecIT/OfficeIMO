using System;
using System.Linq;
using System.Text;

namespace OfficeIMO.Markdown;

public sealed partial class TableBlock {
    private IReadOnlyList<IMarkdownBlock> BuildChildBlocks() {
        var blocks = new List<IMarkdownBlock>();
        var headerCells = HeaderCells;
        for (int i = 0; i < headerCells.Count; i++) {
            for (int j = 0; j < headerCells[i].Blocks.Count; j++) {
                blocks.Add(headerCells[i].Blocks[j]);
            }
        }

        var rowCells = RowCells;
        for (int rowIndex = 0; rowIndex < rowCells.Count; rowIndex++) {
            for (int cellIndex = 0; cellIndex < rowCells[rowIndex].Count; cellIndex++) {
                var cell = rowCells[rowIndex][cellIndex];
                for (int blockIndex = 0; blockIndex < cell.Blocks.Count; blockIndex++) {
                    blocks.Add(cell.Blocks[blockIndex]);
                }
            }
        }

        return blocks;
    }

    private void InvalidateRealizedCellCache() {
        _cachedHeaderCells = null;
        _cachedRowCells = null;
        _cachedCellContentSignature = null;
        _cachedUsesStructuredCells = false;
        _cachedCellColumnCount = -1;
    }

    internal int ComputeContentSignature() {
        unchecked {
            int hash = 17;
            hash = (hash * 31) + Headers.Count;
            for (int i = 0; i < Headers.Count; i++) {
                hash = (hash * 31) + StringComparer.Ordinal.GetHashCode(Headers[i] ?? string.Empty);
            }

            hash = (hash * 31) + Rows.Count;
            for (int rowIndex = 0; rowIndex < Rows.Count; rowIndex++) {
                var row = Rows[rowIndex];
                hash = (hash * 31) + (row?.Count ?? -1);
                if (row == null) {
                    continue;
                }

                for (int cellIndex = 0; cellIndex < row.Count; cellIndex++) {
                    hash = (hash * 31) + StringComparer.Ordinal.GetHashCode(row[cellIndex] ?? string.Empty);
                }
            }

            return hash;
        }
    }

    internal static bool TryConsumeBreakTag(string value, int index, out int consumed) {
        consumed = 0;
        int length = value.Length;
        if (index + 3 >= length) {
            return false;
        }

        if (value[index] != '<') return false;
        if (!IsSpecificLetter(value[index + 1], 'b')) return false;
        if (!IsSpecificLetter(value[index + 2], 'r')) return false;

        int position = index + 3;

        while (position < length && char.IsWhiteSpace(value[position])) {
            position++;
        }

        if (position < length && value[position] == '/') {
            position++;
            while (position < length && char.IsWhiteSpace(value[position])) {
                position++;
            }
        }

        if (position < length && value[position] == '>') {
            consumed = position - index + 1;
            return true;
        }

        return false;
    }

    internal static bool IsSpecificLetter(char value, char expected) {
        return char.ToLowerInvariant(value) == expected;
    }

    internal IReadOnlyList<MarkdownSyntaxNode> BuildSyntaxChildren(MarkdownSourceSpan? span) {
        if (!span.HasValue) {
            return Array.Empty<MarkdownSyntaxNode>();
        }

        var nodes = new List<MarkdownSyntaxNode>();
        int line = span.Value.StartLine;
        int columnCount = GetEffectiveColumnCount();
        var bodyRows = RowCells;

        if (Headers.Count > 0) {
            var headerCells = HeaderCells;
            var headerChildren = BuildTableCellSyntaxChildren(
                PrepareRowCells(Headers, columnCount),
                headerCells,
                new MarkdownSourceSpan(line, line));
            nodes.Add(new MarkdownSyntaxNode(
                MarkdownSyntaxKind.TableHeader,
                MarkdownBlockSyntaxBuilder.GetAggregateSpan(headerChildren) ?? new MarkdownSourceSpan(line, line),
                string.Join(" | ", Headers),
                headerChildren));
            line += 2;
        }

        for (int i = 0; i < Rows.Count; i++) {
            if (line > span.Value.EndLine) {
                break;
            }

            var rowCells = i < bodyRows.Count ? bodyRows[i] : Array.Empty<TableCell>();
            var rowChildren = BuildTableCellSyntaxChildren(
                PrepareRowCells(Rows[i], columnCount),
                rowCells,
                new MarkdownSourceSpan(line, line));
            nodes.Add(new MarkdownSyntaxNode(
                MarkdownSyntaxKind.TableRow,
                MarkdownBlockSyntaxBuilder.GetAggregateSpan(rowChildren) ?? new MarkdownSourceSpan(line, line),
                string.Join(" | ", Rows[i]),
                rowChildren));
            line++;
        }

        return nodes;
    }

    private static IReadOnlyList<MarkdownSyntaxNode> BuildTableCellSyntaxChildren(
        IReadOnlyList<string> rawCells,
        IReadOnlyList<TableCell> structuredCells,
        MarkdownSourceSpan rowSpan) {
        int cellCount = Math.Max(rawCells?.Count ?? 0, structuredCells?.Count ?? 0);
        if (cellCount == 0) {
            return Array.Empty<MarkdownSyntaxNode>();
        }

        var nodes = new List<MarkdownSyntaxNode>(cellCount);
        for (int i = 0; i < cellCount; i++) {
            string literal = rawCells != null && i < rawCells.Count
                ? rawCells[i] ?? string.Empty
                : structuredCells != null && i < structuredCells.Count
                    ? structuredCells[i]?.Markdown ?? string.Empty
                    : string.Empty;
            var cellSpan = structuredCells != null && i < structuredCells.Count
                ? structuredCells[i]?.SourceSpan ?? rowSpan
                : rowSpan;

            IReadOnlyList<MarkdownSyntaxNode> children;
            if (structuredCells != null && i < structuredCells.Count && structuredCells[i]?.SyntaxChildren != null && structuredCells[i]!.SyntaxChildren!.Count > 0) {
                children = structuredCells[i]!.SyntaxChildren!;
            } else if (structuredCells != null && i < structuredCells.Count && structuredCells[i] != null && structuredCells[i].Blocks.Count > 0) {
                var blockNodes = new List<MarkdownSyntaxNode>(structuredCells[i].Blocks.Count);
                for (int blockIndex = 0; blockIndex < structuredCells[i].Blocks.Count; blockIndex++) {
                    blockNodes.Add(MarkdownBlockSyntaxBuilder.BuildBlock(structuredCells[i].Blocks[blockIndex]));
                }
                children = blockNodes;
            } else {
                children = Array.Empty<MarkdownSyntaxNode>();
            }

            nodes.Add(new MarkdownSyntaxNode(
                MarkdownSyntaxKind.TableCell,
                MarkdownBlockSyntaxBuilder.GetAggregateSpan(children) ?? cellSpan,
                literal,
                children,
                structuredCells != null && i < structuredCells.Count ? structuredCells[i] : null));
        }

        return nodes;
    }

    MarkdownSyntaxNode ISyntaxMarkdownBlock.BuildSyntaxNode(MarkdownSourceSpan? span) =>
        new MarkdownSyntaxNode(
            MarkdownSyntaxKind.Table,
            span,
            ((IMarkdownBlock)this).RenderMarkdown(),
            BuildSyntaxChildren(span),
            this);
}
