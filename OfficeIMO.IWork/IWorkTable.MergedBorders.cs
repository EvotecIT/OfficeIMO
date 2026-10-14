using System.Threading;
using OfficeIMO.Drawing;

namespace OfficeIMO.IWork;

public sealed partial class IWorkTable {
    /// <summary>Resolves the outer edges of each merge for destination cells that span multiple grid positions.</summary>
    /// <remarks>Different segments on one edge require an explicit destination approximation. No interior edge is substituted for the perimeter.</remarks>
    internal IEnumerable<(IWorkTableMergeRange Range, OfficeBorderBox Borders, bool IsUniform)> ResolveMergedBorders(CancellationToken token) {
        if (!HasBorderDeclarations) yield break;
        foreach (var merge in MergedRanges) {
            token.ThrowIfCancellationRequested();
            var first = GetBorders(merge.FirstRow, merge.FirstColumn);
            OfficeBorderSide? left = first.Left, top = first.Top;
            OfficeBorderSide? right = GetBorders(merge.FirstRow, merge.LastColumn).Right;
            OfficeBorderSide? bottom = GetBorders(merge.LastRow, merge.FirstColumn).Bottom;
            bool uniform = true;
            for (int row = merge.FirstRow; row <= merge.LastRow; row++) {
                token.ThrowIfCancellationRequested();
                uniform &= Nullable.Equals(left, GetBorders(row, merge.FirstColumn).Left)
                    && Nullable.Equals(right, GetBorders(row, merge.LastColumn).Right);
            }
            for (int column = merge.FirstColumn; column <= merge.LastColumn; column++) {
                token.ThrowIfCancellationRequested();
                uniform &= Nullable.Equals(top, GetBorders(merge.FirstRow, column).Top)
                    && Nullable.Equals(bottom, GetBorders(merge.LastRow, column).Bottom);
            }
            yield return (merge, new OfficeBorderBox(left, top, right, bottom), uniform);
        }
    }
}
