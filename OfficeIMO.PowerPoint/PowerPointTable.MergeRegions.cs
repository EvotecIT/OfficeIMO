using System.Collections.Generic;
using System.Linq;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.PowerPoint;

public partial class PowerPointTable {
    // Row and column mutations work by cell position, so merged ranges are captured
    // first, shifted for the inserted or removed index, and written back afterwards.
    // Without this an anchor keeps a span over cells that no longer continue it.
    private List<(int Row, int Column, int RowSpan, int ColumnSpan)> CaptureMergeRegions() {
        int rows = Rows;
        int columns = Columns;
        var regions = new List<(int Row, int Column, int RowSpan, int ColumnSpan)>();
        foreach ((int row, int column, PowerPointTableCell cell) in EnumerateAnchorCells().ToList()) {
            (int rowSpan, int columnSpan) = cell.Merge;
            rowSpan = Math.Min(rows - row, Math.Max(1, rowSpan));
            columnSpan = Math.Min(columns - column, Math.Max(1, columnSpan));
            if (rowSpan > 1 || columnSpan > 1) regions.Add((row, column, rowSpan, columnSpan));
        }
        return regions;
    }

    /// <summary>Moves an anchor into the next covered cell when its own row is about to be removed.</summary>
    private void PreserveAnchorsForRemovedRow(List<(int Row, int Column, int RowSpan, int ColumnSpan)> regions, int index) {
        A.TableRow[] rows = TableElement.Elements<A.TableRow>().ToArray();
        foreach (var region in regions.Where(_ => _.Row == index && _.RowSpan > 1)) {
            SwapCells(rows[index].Elements<A.TableCell>().ElementAt(region.Column),
                rows[index + 1].Elements<A.TableCell>().ElementAt(region.Column));
        }
    }

    /// <summary>Moves an anchor into the next covered cell when its own column is about to be removed.</summary>
    private void PreserveAnchorsForRemovedColumn(List<(int Row, int Column, int RowSpan, int ColumnSpan)> regions, int index) {
        A.TableRow[] rows = TableElement.Elements<A.TableRow>().ToArray();
        foreach (var region in regions.Where(_ => _.Column == index && _.ColumnSpan > 1)) {
            A.TableCell[] cells = rows[region.Row].Elements<A.TableCell>().ToArray();
            SwapCells(cells[index], cells[index + 1]);
        }
    }

    private static void SwapCells(A.TableCell first, A.TableCell second) {
        var firstParent = first.Parent!;
        var secondParent = second.Parent!;
        var placeholder = new A.TableCell();
        firstParent.ReplaceChild(placeholder, first);
        secondParent.ReplaceChild(first, second);
        firstParent.ReplaceChild(second, placeholder);
    }

    private static void ShiftMergeRegionsForInsertedRow(List<(int Row, int Column, int RowSpan, int ColumnSpan)> regions, int index) {
        for (int i = 0; i < regions.Count; i++) {
            var region = regions[i];
            (region.Row, region.RowSpan) = ShiftSpanForInsertion(region.Row, region.RowSpan, index);
            regions[i] = region;
        }
    }

    private static void ShiftMergeRegionsForInsertedColumn(List<(int Row, int Column, int RowSpan, int ColumnSpan)> regions, int index) {
        for (int i = 0; i < regions.Count; i++) {
            var region = regions[i];
            (region.Column, region.ColumnSpan) = ShiftSpanForInsertion(region.Column, region.ColumnSpan, index);
            regions[i] = region;
        }
    }

    private static void ShiftMergeRegionsForRemovedRow(List<(int Row, int Column, int RowSpan, int ColumnSpan)> regions, int index) {
        for (int i = 0; i < regions.Count; i++) {
            var region = regions[i];
            (region.Row, region.RowSpan) = ShiftSpanForRemoval(region.Row, region.RowSpan, index);
            regions[i] = region;
        }
    }

    private static void ShiftMergeRegionsForRemovedColumn(List<(int Row, int Column, int RowSpan, int ColumnSpan)> regions, int index) {
        for (int i = 0; i < regions.Count; i++) {
            var region = regions[i];
            (region.Column, region.ColumnSpan) = ShiftSpanForRemoval(region.Column, region.ColumnSpan, index);
            regions[i] = region;
        }
    }

    // An index inserted strictly inside a span widens it; at or before the start it moves the span.
    private static (int Start, int Span) ShiftSpanForInsertion(int start, int span, int index) =>
        index <= start ? (start + 1, span)
        : index < start + span ? (start, span + 1)
        : (start, span);

    private static (int Start, int Span) ShiftSpanForRemoval(int start, int span, int index) =>
        index < start ? (start - 1, span)
        : index < start + span ? (start, span - 1)
        : (start, span);

    private void ApplyMergeRegions(List<(int Row, int Column, int RowSpan, int ColumnSpan)> regions) {
        A.TableCell[][] cells = TableElement.Elements<A.TableRow>()
            .Select(_ => _.Elements<A.TableCell>().ToArray())
            .ToArray();
        foreach (A.TableCell cell in cells.SelectMany(_ => _)) {
            cell.RowSpan = null;
            cell.GridSpan = null;
            cell.HorizontalMerge = null;
            cell.VerticalMerge = null;
        }

        foreach (var region in regions) {
            if (region.RowSpan < 1 || region.ColumnSpan < 1 || region.RowSpan == 1 && region.ColumnSpan == 1) continue;
            for (int row = region.Row; row < region.Row + region.RowSpan; row++) {
                for (int column = region.Column; column < region.Column + region.ColumnSpan; column++) {
                    A.TableCell cell = cells[row][column];
                    if (row == region.Row && column == region.Column) {
                        cell.RowSpan = region.RowSpan > 1 ? region.RowSpan : null;
                        cell.GridSpan = region.ColumnSpan > 1 ? region.ColumnSpan : null;
                        continue;
                    }
                    cell.HorizontalMerge = column > region.Column ? true : (bool?)null;
                    cell.VerticalMerge = row > region.Row ? true : (bool?)null;
                }
            }
        }
    }
}
