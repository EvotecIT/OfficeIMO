using System.Threading;

namespace OfficeIMO.IWork.Internal;

/// <summary>Assesses only decimal tabs that survive paragraph overrides and table merging.</summary>
internal static class IWorkDecimalTabLayout {
    internal static IEnumerable<string?> Symbols(IWorkTextContent content, IWorkParagraphStyle? defaults = null, CancellationToken cancellationToken = default) {
        foreach (IWorkTextParagraph paragraph in content.Paragraphs) {
            cancellationToken.ThrowIfCancellationRequested();
            if (HasDecimalStop(paragraph.Style.TabStops ?? defaults?.TabStops))
                yield return paragraph.Style.DecimalTabSymbolDeclared ? paragraph.Style.DecimalTabSymbol : defaults?.DecimalTabSymbol;
        }
    }

    internal static IEnumerable<string?> Symbols(IWorkTable table, CancellationToken cancellationToken) {
        // Most tables have no decimal stops. Do not enumerate their sparse grid for this assessment.
        if (!new[] { table.TextStyles.Body, table.TextStyles.HeaderRow, table.TextStyles.HeaderColumn, table.TextStyles.FooterRow }
            .Concat(table.Cells.Select(cell => cell.ParagraphStyle))
            .Concat(table.Cells.Where(cell => cell.RichText != null).SelectMany(cell => cell.RichText!.Paragraphs).Select(paragraph => paragraph.Style))
            .Any(style => { cancellationToken.ThrowIfCancellationRequested(); return HasDecimalStop(style?.TabStops); })) yield break;

        IWorkTableMergeRange[] merges = table.MergedRanges.OrderBy(range => range.FirstRow).ToArray();
        var active = new SortedDictionary<int, IWorkTableMergeRange>();
        int nextMerge = 0;
        for (int row = 1; row <= table.RowCount; row++) {
            cancellationToken.ThrowIfCancellationRequested();
            foreach (int column in active.Where(pair => pair.Value.LastRow < row).Select(pair => pair.Key).ToArray()) active.Remove(column);
            while (nextMerge < merges.Length && merges[nextMerge].FirstRow == row) {
                IWorkTableMergeRange merge = merges[nextMerge++];
                active.Add(merge.FirstColumn, merge);
            }
            IWorkTableMergeRange[] rowMerges = active.Values.ToArray();
            int mergeIndex = 0;
            for (int column = 1; column <= table.ColumnCount; column++) {
                cancellationToken.ThrowIfCancellationRequested();
                while (mergeIndex < rowMerges.Length && rowMerges[mergeIndex].LastColumn < column) mergeIndex++;
                if (mergeIndex < rowMerges.Length && rowMerges[mergeIndex].FirstColumn <= column
                    && (rowMerges[mergeIndex].FirstRow != row || rowMerges[mergeIndex].FirstColumn != column)) continue;
                IWorkParagraphStyle? defaults = table.GetParagraphStyle(row, column);
                if (table.GetCell(row, column)?.RichText is { Paragraphs.Count: > 0 } richText) {
                    foreach (string? symbol in Symbols(richText, defaults, cancellationToken)) yield return symbol;
                } else if (HasDecimalStop(defaults?.TabStops)) yield return defaults!.DecimalTabSymbol;
            }
        }
    }

    private static bool HasDecimalStop(IReadOnlyList<IWorkTabStop>? stops) =>
        stops?.Any(stop => stop.Alignment == IWorkTabAlignment.Decimal) == true;
}
