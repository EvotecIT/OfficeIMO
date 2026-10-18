using System;
using System.Collections.Generic;

namespace PSWriteOffice.Services.Table;

internal static class OfficeTableGridPlanner
{
    public static IReadOnlyList<OfficeTableCellPlacement> Plan(
        IReadOnlyList<IReadOnlyList<OfficeTableCellSpec>> rows,
        out int columnCount)
    {
        if (rows == null)
        {
            throw new ArgumentNullException(nameof(rows));
        }

        var placements = new List<OfficeTableCellPlacement>();
        var activeRowSpans = new Dictionary<int, int>();
        columnCount = 0;

        for (var rowIndex = 0; rowIndex < rows.Count; rowIndex++)
        {
            var occupiedColumns = new HashSet<int>(activeRowSpans.Keys);
            var futureRowSpans = new Dictionary<int, int>();
            var columnIndex = 0;

            foreach (var cell in rows[rowIndex])
            {
                while (IsRangeOccupied(occupiedColumns, columnIndex, cell.ColumnSpan))
                {
                    columnIndex++;
                }

                placements.Add(new OfficeTableCellPlacement(rowIndex, columnIndex, cell));

                for (var offset = 0; offset < cell.ColumnSpan; offset++)
                {
                    var occupiedColumn = columnIndex + offset;
                    occupiedColumns.Add(occupiedColumn);
                    if (cell.RowSpan > 1)
                    {
                        futureRowSpans[occupiedColumn] = Math.Max(
                            futureRowSpans.TryGetValue(occupiedColumn, out var existing) ? existing : 0,
                            cell.RowSpan - 1);
                    }
                }

                columnIndex += cell.ColumnSpan;
            }

            foreach (var entry in activeRowSpans)
            {
                var remainingRows = entry.Value - 1;
                if (remainingRows > 0)
                {
                    futureRowSpans[entry.Key] = Math.Max(
                        futureRowSpans.TryGetValue(entry.Key, out var existing) ? existing : 0,
                        remainingRows);
                }
            }

            if (occupiedColumns.Count > 0)
            {
                foreach (var occupiedColumn in occupiedColumns)
                {
                    columnCount = Math.Max(columnCount, occupiedColumn + 1);
                }
            }

            activeRowSpans = futureRowSpans;
        }

        return placements;
    }

    private static bool IsRangeOccupied(HashSet<int> occupiedColumns, int columnIndex, int columnSpan)
    {
        for (var offset = 0; offset < columnSpan; offset++)
        {
            if (occupiedColumns.Contains(columnIndex + offset))
            {
                return true;
            }
        }

        return false;
    }
}
