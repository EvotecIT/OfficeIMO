using System.Threading;
using OfficeIMO.IWork;

namespace OfficeIMO.Word.IWork;

public static partial class WordIWorkConverter {
    /// <summary>Maps native physical row boxes to Word's authored height declarations.</summary>
    private static int?[] WordTableRowHeights(IWorkTable source, CancellationToken cancellationToken) {
        var heights = new int?[source.RowCount];
        for (int row = 1; row <= source.RowCount; row++) {
            cancellationToken.ThrowIfCancellationRequested();
            if (source.GetRowHeight(row) is double height) heights[row - 1] = ToSignedTwips(height);
        }
        // Automatic rows retain their separately qualified minimum-height contract.
        if (source.AutoResizeRows == true) return heights;

        // Word adds the row's largest active bottom margin to an exact trHeight.
        // The adapter creates TableGrid tables, whose inherited top/bottom margins are zero.
        // Horizontal continuations disappear when Word normalizes the merges; vertical
        // continuations retain their own margins and still affect their physical row.
        // Reconstruction validation bounds this grid before either caller reaches this helper.
        var horizontalContinuations = new HashSet<(int Row, int Column)>();
        foreach (IWorkTableMergeRange merge in source.MergedRanges) {
            for (int row = merge.FirstRow; row <= merge.LastRow; row++) {
                cancellationToken.ThrowIfCancellationRequested();
                for (int column = merge.FirstColumn + 1; column <= merge.LastColumn; column++)
                    horizontalContinuations.Add((row, column));
            }
        }
        for (int row = 1; row <= source.RowCount; row++) {
            cancellationToken.ThrowIfCancellationRequested();
            int largestBottomMargin = 0;
            for (int column = 1; column <= source.ColumnCount; column++) {
                if (horizontalContinuations.Contains((row, column))) continue;
                if (source.GetPadding(row, column) is { } padding)
                    largestBottomMargin = Math.Max(largestBottomMargin, ToSignedTwips(padding.BottomPoints));
            }
            if (heights[row - 1] is int height) heights[row - 1] = height - largestBottomMargin;
        }
        return heights;
    }
}
