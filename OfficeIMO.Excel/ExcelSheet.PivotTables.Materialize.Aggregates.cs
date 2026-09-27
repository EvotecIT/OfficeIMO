using DocumentFormat.OpenXml.Spreadsheet;
using System.Threading;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        private static Dictionary<(int Row, int Column), ExcelPivotAggregateAccumulator[]> AggregateMaterializedHierarchy(
            ExcelSheet source, int firstRow, int firstColumn, int lastRow, PivotHierarchyAxis rows, PivotHierarchyAxis columns,
            DataField[] measures, int limit, CancellationToken token) {
            var groups = new Dictionary<(int Row, int Column), ExcelPivotAggregateAccumulator[]>();
            var values = new object?[measures.Length];
            var errors = new bool[measures.Length];
            int leafGroups = 0;
            for (int row = firstRow + 1; row <= lastRow; row++) {
                token.ThrowIfCancellationRequested();
                for (int measure = 0; measure < measures.Length; measure++) {
                    var cell = source.TryGetExistingCell(row, firstColumn + (int)measures[measure].Field!.Value);
                    values[measure] = source.GetCellValueSnapshot(cell).Value;
                    errors[measure] = cell?.DataType?.Value == DocumentFormat.OpenXml.Spreadsheet.CellValues.Error;
                }
                var rowLeaf = rows.Leaf(row - firstRow - 1);
                var columnLeaf = columns.Leaf(row - firstRow - 1);
                for (var rowNode = rowLeaf; rowNode != null; rowNode = rowNode.Parent) {
                    if (!rows.Displays(rowNode)) continue;
                    for (var columnNode = columnLeaf; columnNode != null; columnNode = columnNode.Parent) {
                        if (!columns.Displays(columnNode)) continue;
                        if (!groups.TryGetValue((rowNode.Id, columnNode.Id), out var group)) {
                            bool leaf = ReferenceEquals(rowNode, rowLeaf) && ReferenceEquals(columnNode, columnLeaf);
                            if (leaf && leafGroups >= 100_000 / measures.Length)
                                throw new InvalidOperationException("The pivot exceeds 100,000 observed leaf aggregate states across its measures.");
                            if (groups.Count >= limit / measures.Length)
                                throw new InvalidOperationException("The pivot aggregate states exceed the materialization budget.");
                            group = Enumerable.Range(0, measures.Length).Select(_ => new ExcelPivotAggregateAccumulator()).ToArray();
                            groups.Add((rowNode.Id, columnNode.Id), group);
                            if (leaf) leafGroups++;
                        }
                        for (int measure = 0; measure < measures.Length; measure++) group[measure].Add(values[measure], errors[measure]);
                    }
                }
            }
            return groups;
        }
    }
}
