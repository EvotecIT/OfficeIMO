using DocumentFormat.OpenXml.Spreadsheet;
using System.Threading;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        private sealed class PivotHierarchyNode {
            internal int Id, Key, Depth;
            internal PivotHierarchyNode? Parent;
            internal Dictionary<int, PivotHierarchyNode>? ChildrenByKey;
            internal List<PivotHierarchyNode>? Children;
        }

        private readonly struct PivotHierarchyEntry {
            internal PivotHierarchyEntry(PivotHierarchyNode node, int measure, ItemValues type) { Node = node; Measure = measure; Type = type; }
            internal PivotHierarchyNode Node { get; }
            internal int Measure { get; }
            internal ItemValues Type { get; }
        }

        private sealed class PivotHierarchyAxis {
            internal PivotMaterializationAxis Layout = null!;
            internal PivotHierarchyNode Root = new() { Key = -1 };
            internal PivotHierarchyNode[]? Leaves;
            internal bool[] Subtotals = Array.Empty<bool>();
            internal bool GrandTotal;
            internal List<PivotHierarchyEntry> Entries = new();
            internal PivotHierarchyNode Leaf(int record) => Leaves == null ? Root : Leaves[record];
            internal bool Displays(PivotHierarchyNode node) => node.Depth == Layout.RealFields.Length
                || (node.Depth == 0 ? GrandTotal : Subtotals[node.Depth - 1]);
        }

        private static bool MaterializedAutomaticSubtotal(PivotField field) => field.DefaultSubtotal?.Value != false
            || field.SumSubtotal?.Value == true || field.CountASubtotal?.Value == true || field.AverageSubTotal?.Value == true
            || field.MaxSubtotal?.Value == true || field.MinSubtotal?.Value == true || field.ApplyProductInSubtotal?.Value == true
            || field.CountSubtotal?.Value == true || field.ApplyStandardDeviationInSubtotal?.Value == true
            || field.ApplyStandardDeviationPInSubtotal?.Value == true || field.ApplyVarianceInSubtotal?.Value == true
            || field.ApplyVariancePInSubtotal?.Value == true;

        // Grand totals add at most one prefix per axis; keeping them outside the
        // input budget preserves flat-view workloads and bounds actual updates
        // by at most four times the budget, including both grand-total axes.
        private static int MaterializedInputLevels(PivotMaterializationAxis axis, PivotField[] fields) => axis.RealFields.Length == 0 ? 1
            : 1 + axis.RealFields.Take(axis.RealFields.Length - 1).Count(field => MaterializedAutomaticSubtotal(fields[field]));

        private static PivotHierarchyAxis BuildMaterializedHierarchyAxis(ExcelSheet source, int firstRow, int lastRow, int firstColumn,
            IReadOnlyList<PivotFieldValues> maps, PivotMaterializationAxis layout, PivotField[] fields,
            IReadOnlyDictionary<int, PivotNumericGrouping> groupings,
            IReadOnlyDictionary<int, PivotDateGrouping> dateGroupings, bool[] includedRows,
            int measures, bool total, CancellationToken token) {
            var result = new PivotHierarchyAxis { Layout = layout, GrandTotal = total,
                Subtotals = layout.RealFields.Select(field => MaterializedAutomaticSubtotal(fields[field])).ToArray() };
            if (layout.RealFields.Length > 0) {
                result.Leaves = new PivotHierarchyNode[lastRow - firstRow];
                var indices = layout.RealFields.Select(field => IndexPivotMaterializationKeys(maps[field])).ToArray();
                int nodeId = 0;
                for (int row = firstRow + 1; row <= lastRow; row++) {
                    token.ThrowIfCancellationRequested();
                    if (!includedRows[row - firstRow - 1]) continue;
                    var node = result.Root;
                    for (int depth = 0; depth < layout.RealFields.Length; depth++) {
                        int field = layout.RealFields[depth];
                        int sourceField = dateGroupings.TryGetValue(field, out var date) ? date.SourceField : field;
                        int key = indices[depth][MaterializedPivotAxisKey(source, row, firstColumn + sourceField, field, groupings, dateGroupings)];
                        node.ChildrenByKey ??= new Dictionary<int, PivotHierarchyNode>();
                        if (!node.ChildrenByKey.TryGetValue(key, out var child)) {
                            child = new PivotHierarchyNode { Id = ++nodeId, Key = key, Depth = depth + 1, Parent = node };
                            node.ChildrenByKey.Add(key, child);
                            (node.Children ??= new List<PivotHierarchyNode>()).Add(child);
                        }
                        node = child;
                    }
                    result.Leaves[row - firstRow - 1] = node;
                }
            }
            void Sort(PivotHierarchyNode node) {
                token.ThrowIfCancellationRequested();
                if (node.Children == null) return;
                node.Children.Sort((left, right) => left.Key.CompareTo(right.Key));
                foreach (var child in node.Children) Sort(child);
                node.ChildrenByKey = null;
            }
            Sort(result.Root);
            void Add(PivotHierarchyNode node, int measure, ItemValues type) {
                if (result.Entries.Count >= 100_001) throw new InvalidOperationException("The generated pivot axis exceeds 100,001 saved items.");
                result.Entries.Add(new PivotHierarchyEntry(node, measure, type));
            }
            void Visit(PivotHierarchyNode node, int position, int measure, bool hasMeasure) {
                token.ThrowIfCancellationRequested();
                if (position == layout.Fields.Length) { Add(node, measure, ItemValues.Data); return; }
                if (layout.Fields[position] == -2) {
                    for (int value = 0; value < measures; value++) Visit(node, position + 1, value, true);
                } else {
                    foreach (var child in node.Children!) {
                        Visit(child, position + 1, measure, hasMeasure);
                        if (child.Depth < layout.RealFields.Length && result.Subtotals[child.Depth - 1]) {
                            if (layout.HasValues && !hasMeasure) for (int value = 0; value < measures; value++) Add(child, value, ItemValues.Default);
                            else Add(child, measure, ItemValues.Default);
                        }
                    }
                }
            }
            Visit(result.Root, 0, 0, false);
            if (total) for (int measure = 0; measure < (layout.HasValues ? measures : 1); measure++) Add(result.Root, measure, ItemValues.Grand);
            return result;
        }

        private static RowItem CreateMaterializedHierarchyItem(PivotHierarchyAxis axis, PivotHierarchyEntry entry) {
            var item = new RowItem();
            if (axis.Layout.HasValues) item.Index = (uint)entry.Measure;
            if (entry.Type == ItemValues.Grand) {
                item.ItemType = ItemValues.Grand;
                item.AppendChild(new MemberPropertyIndex { Val = 0 });
                return item;
            }
            if (entry.Type == ItemValues.Default) item.ItemType = ItemValues.Default;
            var keys = MaterializedHierarchyKeys(entry.Node);
            int depth = 0;
            foreach (int field in axis.Layout.Fields) {
                if (entry.Type == ItemValues.Default && depth == keys.Length) break;
                item.AppendChild(new MemberPropertyIndex { Val = field == -2 ? entry.Measure : keys[depth++] });
            }
            return item;
        }

        private static int[] MaterializedHierarchyKeys(PivotHierarchyNode node) {
            var keys = new int[node.Depth];
            for (var current = node; current.Parent != null; current = current.Parent) keys[current.Depth - 1] = current.Key;
            return keys;
        }
    }
}
