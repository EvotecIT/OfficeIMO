using System;
using System.Linq;
using System.Text;

namespace OfficeIMO.Markdown;

public sealed partial class TableBlock {
    private IReadOnlyList<IMarkdownBlock> BuildChildBlocks() {
        var blocks = new List<IMarkdownBlock>();
        var headerCells = HeaderCells;
        for (int i = 0; i < headerCells.Count; i++) {
            for (int j = 0; j < headerCells[i].ChildBlocks.Count; j++) {
                blocks.Add(headerCells[i].ChildBlocks[j]);
            }
        }

        var rowCells = RowCells;
        for (int rowIndex = 0; rowIndex < rowCells.Count; rowIndex++) {
            for (int cellIndex = 0; cellIndex < rowCells[rowIndex].Count; cellIndex++) {
                var cell = rowCells[rowIndex][cellIndex];
                for (int blockIndex = 0; blockIndex < cell.ChildBlocks.Count; blockIndex++) {
                    blocks.Add(cell.ChildBlocks[blockIndex]);
                }
            }
        }

        return blocks;
    }

    private void InvalidateRealizedCellCache() {
        _cachedHeaderCells = null;
        _cachedRowCells = null;
        _cachedHeaderRow = null;
        _cachedBodyRows = null;
        _cachedCellContentSignature = null;
        _cachedUsesStructuredCells = false;
        _cachedCellColumnCount = -1;
    }

    private TableRow? GetOrBuildHeaderRow() {
        if (Headers.Count == 0) {
            return null;
        }

        return _cachedHeaderRow ??= new TableRow(HeaderCells, isHeader: true, rowIndex: -1);
    }

    private IReadOnlyList<TableRow> GetOrBuildBodyRows() {
        if (_cachedBodyRows != null) {
            return _cachedBodyRows;
        }

        var rowCells = RowCells;
        if (rowCells.Count == 0) {
            _cachedBodyRows = Array.Empty<TableRow>();
            return _cachedBodyRows;
        }

        var rows = new TableRow[rowCells.Count];
        for (int i = 0; i < rowCells.Count; i++) {
            rows[i] = new TableRow(rowCells[i], isHeader: false, rowIndex: i);
        }

        _cachedBodyRows = rows;
        return _cachedBodyRows;
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
        var bodyRowOwners = BodyRows;

        if (Headers.Count > 0) {
            var headerCells = HeaderCells;
            var headerRow = HeaderRow;
            var headerChildren = BuildTableCellSyntaxChildren(
                PrepareRowCells(Headers, columnCount),
                headerCells,
                new MarkdownSourceSpan(line, line));
            nodes.Add(new MarkdownSyntaxNode(
                MarkdownSyntaxKind.TableHeader,
                MarkdownBlockSyntaxBuilder.GetAggregateSpan(headerChildren) ?? new MarkdownSourceSpan(line, line),
                string.Join(" | ", Headers),
                headerChildren,
                headerRow));
            nodes.Add(new MarkdownSyntaxNode(
                MarkdownSyntaxKind.TableAlignmentRow,
                GetAlignmentRowSourceSpan(span, line + 1),
                BuildAlignmentRowLiteral(columnCount)));
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
                rowChildren,
                i < bodyRowOwners.Count ? bodyRowOwners[i] : null));
            line++;
        }

        return nodes;
    }

    private MarkdownSourceSpan? GetAlignmentRowSourceSpan(MarkdownSourceSpan? tableSpan, int line) {
        if (AlignmentRowSourceSpan.HasValue) {
            return AlignmentRowSourceSpan;
        }

        if (!tableSpan.HasValue) {
            return null;
        }

        return new MarkdownSourceSpan(line, line);
    }

    private string BuildAlignmentRowLiteral(int columnCount) {
        if (columnCount <= 0) {
            return string.Empty;
        }

        var cells = new string[columnCount];
        for (int i = 0; i < columnCount; i++) {
            cells[i] = GetAlignment(i) switch {
                ColumnAlignment.Left => ":---",
                ColumnAlignment.Center => ":---:",
                ColumnAlignment.Right => "---:",
                _ => "---"
            };
        }

        return string.Join(" | ", cells);
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

            IReadOnlyList<MarkdownSyntaxNode> children = structuredCells != null && i < structuredCells.Count && structuredCells[i] != null
                ? MarkdownBlockSyntaxBuilder.GetOwnedSyntaxChildrenOrBuild(structuredCells[i]!)
                : Array.Empty<MarkdownSyntaxNode>();

            nodes.Add(new MarkdownSyntaxNode(
                MarkdownSyntaxKind.TableCell,
                cellSpan,
                literal,
                children,
                structuredCells != null && i < structuredCells.Count ? structuredCells[i] : null));
        }

        return nodes;
    }

    IReadOnlyList<MarkdownSyntaxNode>? ISyntaxChildrenMarkdownBlock.ProvidedSyntaxChildren => null;

    IReadOnlyList<MarkdownSyntaxNode> IOwnedSyntaxChildrenMarkdownBlock.BuildOwnedSyntaxChildren() =>
        BuildSyntaxChildren(SourceSpan);

    MarkdownSyntaxNode ISyntaxMarkdownBlock.BuildSyntaxNode(MarkdownSourceSpan? span) {
        var children = SourceSpan.HasValue || !span.HasValue
            ? MarkdownBlockSyntaxBuilder.GetOwnedSyntaxChildrenOrBuild(this)
            : BuildSyntaxChildren(span);

        return new MarkdownSyntaxNode(
            MarkdownSyntaxKind.Table,
            span,
            ((IMarkdownBlock)this).RenderMarkdown(),
            children,
            this);
    }
}
