using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        private static Dictionary<int, Dictionary<PivotFieldValue, string>> ReadPivotMaterializationCaptions(
            CacheField[] cacheFields, PivotField[] pivotFields, IReadOnlyList<PivotFieldValues> maps,
            IEnumerable<int> activeFields,
            IReadOnlyDictionary<int, PivotNumericGrouping> numericGroupings,
            IReadOnlyDictionary<int, PivotDateGrouping> dateGroupings,
            IReadOnlyDictionary<int, PivotManualGrouping> manualGroupings) {
            var result = new Dictionary<int, Dictionary<PivotFieldValue, string>>();
            foreach (int field in activeFields.Distinct()) {
                // Manual group labels already use the saved pivot-item captions as keys.
                if (manualGroupings.ContainsKey(field)) continue;
                var items = pivotFields[field].Items?.Elements<Item>();
                if (items == null) continue;
                var renamed = items.Where(item => item.ItemType == null || item.ItemType.Value == ItemValues.Data)
                    .Select(item => (Item: item, Caption: PivotLookupAttribute(item, "n")))
                    .Where(entry => entry.Caption.Length != 0).ToArray();
                if (renamed.Length == 0) continue;
                OpenXmlElement[] shared = dateGroupings.TryGetValue(field, out var date) ? date.SavedItems
                    : numericGroupings.TryGetValue(field, out var numeric) ? numeric.SavedItems
                    : cacheFields[field].SharedItems?.ChildElements.ToArray() ?? Array.Empty<OpenXmlElement>();
                var current = new HashSet<PivotFieldValue>(maps[field].Items);
                var captions = new Dictionary<PivotFieldValue, string>();
                foreach (var entry in renamed) {
                    PivotFieldValue key = OriginalPivotMaterializationKey(entry.Item, shared, false);
                    if (current.Contains(key)) captions.Add(key, entry.Caption);
                }
                if (captions.Count > 0) result.Add(field, captions);
            }
            return result;
        }
    }
}
