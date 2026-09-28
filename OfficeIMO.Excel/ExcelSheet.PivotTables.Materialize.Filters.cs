using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Threading;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        private sealed class PivotMaterializationVisibility {
            internal bool[] IncludedRows = Array.Empty<bool>();
            internal readonly Dictionary<int, HashSet<PivotFieldValue>> Hidden = new();
            internal readonly Dictionary<int, PivotFieldValue> SelectedPages = new();
        }

        private PivotMaterializationVisibility BuildPivotMaterializationVisibility(
            ExcelSheet source, CacheField[] cacheFields, PivotField[] pivotFields, PageField[] pages,
            IReadOnlyList<PivotFieldValues> maps, IReadOnlyList<int> axisFields,
            IReadOnlyDictionary<int, PivotNumericGrouping> groupings,
            IReadOnlyDictionary<int, PivotDateGrouping> dateGroupings,
            IReadOnlyDictionary<int, PivotManualGrouping> manualGroupings,
            int firstRow, int lastRow, int firstColumn, int limit, CancellationToken token) {
            var result = new PivotMaterializationVisibility { IncludedRows = new bool[lastRow - firstRow] };
            var filtered = new HashSet<int>();
            foreach (int field in axisFields.Concat(pages.Select(page => page.Field!.Value)).Distinct()) {
                Item[] items = pivotFields[field].Items?.Elements<Item>().ToArray() ?? Array.Empty<Item>();
                bool hidden = items.Any(item => item.Hidden?.Value == true);
                PageField? page = pages.FirstOrDefault(candidate => candidate.Field?.Value == field);
                if (!hidden && page?.Item == null) continue;
                if (items.Length > 100_001 || cacheFields[field].SharedItems?.ChildElements.Count > 100_000)
                    throw new NotSupportedException("The filtered pivot field exceeds materialization limits.");
                OpenXmlElement[] shared = manualGroupings.TryGetValue(field, out var manual) ? manual.SavedItems
                    : dateGroupings.TryGetValue(field, out var date) ? date.SavedItems
                    : groupings.TryGetValue(field, out var numeric) ? numeric.SavedItems
                    : cacheFields[field].SharedItems?.ChildElements.ToArray()
                    ?? Array.Empty<OpenXmlElement>();
                if (shared.Length == 0 || items.Length == 0)
                    throw new NotSupportedException("The filtered pivot field has no saved item mapping.");
                var hiddenKeys = new HashSet<PivotFieldValue>();
                foreach (Item item in items) {
                    if (item.Hidden?.Value != true) continue;
                    hiddenKeys.Add(OriginalPivotMaterializationKey(item, shared));
                }
                if (hiddenKeys.Count > 0 && pivotFields[field].IncludeNewItemsInFilter?.Value != true) {
                    var known = new HashSet<PivotFieldValue>(items.Where(item => item.ItemType == null || item.ItemType.Value == ItemValues.Data)
                        .Select(item => OriginalPivotMaterializationKey(item, shared)));
                    foreach (PivotFieldValue key in maps[field].Items)
                        if (!known.Contains(key)) hiddenKeys.Add(key);
                }
                if (hiddenKeys.Count > 0) result.Hidden.Add(field, hiddenKeys);
                if (page?.Item != null) {
                    if (page.Item.Value >= items.Length)
                        throw new NotSupportedException("The selected page item is outside the saved pivot field.");
                    Item selected = items[page.Item.Value];
                    if (selected.ItemType?.Value != ItemValues.Default) {
                        PivotFieldValue key = OriginalPivotMaterializationKey(selected, shared);
                        if (hiddenKeys.Contains(key))
                            throw new NotSupportedException("The selected page item is hidden.");
                        result.SelectedPages.Add(field, key);
                    }
                }
                filtered.Add(field);
            }
            if ((long)(lastRow - firstRow) * filtered.Count > limit)
                throw new InvalidOperationException("The filtered pivot source exceeds the materialization work budget.");
            int visible = 0;
            for (int row = firstRow + 1; row <= lastRow; row++) {
                token.ThrowIfCancellationRequested();
                bool include = true;
                foreach (int field in filtered) {
                    int sourceField = dateGroupings.TryGetValue(field, out var date) ? date.SourceField
                        : manualGroupings.TryGetValue(field, out var manual) ? manual.SourceField : field;
                    PivotFieldValue key = MaterializedPivotAxisKey(source, row, firstColumn + sourceField,
                        field, groupings, dateGroupings, manualGroupings);
                    if (result.Hidden.TryGetValue(field, out var hiddenKeys) && hiddenKeys.Contains(key)
                        || result.SelectedPages.TryGetValue(field, out var selected) && !selected.Equals(key)) {
                        include = false;
                        break;
                    }
                }
                result.IncludedRows[row - firstRow - 1] = include;
                if (include) visible++;
            }
            if (visible == 0 && filtered.Count > 0)
                throw new NotSupportedException("The pivot filters select no source records for materialization.");
            foreach (var pair in result.SelectedPages)
                if (!maps[pair.Key].Items.Contains(pair.Value))
                    throw new NotSupportedException("The selected page item is absent from the current source.");
            return result;
        }

        private static PivotFieldValue OriginalPivotMaterializationKey(Item item, OpenXmlElement[] shared) {
            if (item.Index == null || item.Index.Value >= shared.Length)
                throw new NotSupportedException("The filtered pivot item has no valid shared key.");
            string caption = PivotLookupAttribute(item, "n");
            if (caption.Length != 0) return PivotFieldValue.FromText(caption);
            return OriginalPivotMaterializationKey(shared[item.Index.Value]);
        }

        private static PivotFieldValue OriginalPivotMaterializationKey(OpenXmlElement item) => item switch {
                StringItem text => PivotFieldValue.FromText(text.Val?.Value ?? string.Empty),
                NumberItem number when number.Val != null => PivotFieldValue.FromNumber(number.Val.Value),
                BooleanItem boolean when boolean.Val != null => PivotFieldValue.FromBoolean(boolean.Val.Value),
                DateTimeItem date when date.Val != null => PivotFieldValue.FromDate(date.Val.Value),
                ErrorItem error => PivotFieldValue.FromError(error.Val?.Value ?? string.Empty),
                MissingItem => PivotFieldValue.Blank(),
                _ => throw new NotSupportedException("The filtered pivot key type is not supported.")
            };

        private static void NormalizeMaterializedPageFields(PivotTableDefinition definition,
            IReadOnlyList<PivotFieldValues> maps, PivotMaterializationVisibility visibility) {
            var fields = definition.PivotFields!.Elements<PivotField>().ToArray();
            foreach (PageField page in definition.PageFields?.Elements<PageField>() ?? Enumerable.Empty<PageField>()) {
                int field = page.Field!.Value;
                fields[field].Items = CreateMaterializedFilteredItems(maps[field],
                    visibility.Hidden.TryGetValue(field, out var hidden) ? hidden : null, true, false);
                if (visibility.SelectedPages.TryGetValue(field, out var selected)) {
                    int index = maps[field].Items.ToList().FindIndex(item => item.Equals(selected));
                    if (index < 0) throw new NotSupportedException("The selected page item is absent from the current source.");
                    page.Item = (uint)index;
                } else page.Item = null;
            }
        }

        private static Items CreateMaterializedFilteredItems(PivotFieldValues map,
            HashSet<PivotFieldValue>? hidden, bool includeDefault, bool subtotal) {
            var items = new Items { Count = (uint)(map.Items.Count + (includeDefault || subtotal ? 1 : 0)) };
            for (int index = 0; index < map.Items.Count; index++) {
                var item = new Item { Index = (uint)index };
                if (hidden?.Contains(map.Items[index]) == true) item.Hidden = true;
                items.AppendChild(item);
            }
            if (includeDefault || subtotal) items.AppendChild(new Item { ItemType = ItemValues.Default });
            return items;
        }
    }
}
