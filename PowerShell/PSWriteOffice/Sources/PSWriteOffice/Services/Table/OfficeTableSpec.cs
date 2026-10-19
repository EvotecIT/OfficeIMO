using System;
using System.Collections.Generic;

namespace PSWriteOffice.Services.Table;

internal sealed class OfficeTableSpec
{
    public OfficeTableSpec(IReadOnlyList<IReadOnlyList<OfficeTableCellSpec>> rows, int? headerRowIndex)
    {
        Rows = rows ?? throw new ArgumentNullException(nameof(rows));
        HeaderRowIndex = headerRowIndex;
        Placements = OfficeTableGridPlanner.Plan(rows, out var columnCount);
        ColumnCount = columnCount;
        ValidateSpanBounds(Placements, RowCount);
    }

    public IReadOnlyList<IReadOnlyList<OfficeTableCellSpec>> Rows { get; }

    public bool HasHeader => HeaderRowIndex.HasValue;

    public int? HeaderRowIndex { get; }

    public IReadOnlyList<OfficeTableCellPlacement> Placements { get; }

    public int RowCount => Rows.Count;

    public int ColumnCount { get; }

    private static void ValidateSpanBounds(IReadOnlyList<OfficeTableCellPlacement> placements, int rowCount)
    {
        foreach (var placement in placements)
        {
            if (placement.RowIndex + placement.Cell.RowSpan > rowCount)
            {
                throw new ArgumentOutOfRangeException(
                    nameof(placement.Cell.RowSpan),
                    "Row span cannot extend past the last table row.");
            }
        }
    }
}

internal sealed class OfficeTableCellPlacement
{
    public OfficeTableCellPlacement(int rowIndex, int columnIndex, OfficeTableCellSpec cell)
    {
        RowIndex = rowIndex;
        ColumnIndex = columnIndex;
        Cell = cell ?? throw new ArgumentNullException(nameof(cell));
    }

    public int RowIndex { get; }

    public int ColumnIndex { get; }

    public OfficeTableCellSpec Cell { get; }
}
