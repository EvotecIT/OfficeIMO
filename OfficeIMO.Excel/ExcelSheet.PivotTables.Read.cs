using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Text;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {

        private static Dictionary<uint, PivotCacheDefinition> BuildPivotCacheMap(WorkbookPart workbookPart) {
            var map = new Dictionary<uint, PivotCacheDefinition>();
            var workbook = workbookPart.Workbook ?? throw new InvalidOperationException("Workbook is missing.");
            var pivotCaches = workbook.PivotCaches;
            if (pivotCaches == null) return map;
            foreach (var cache in pivotCaches.Elements<PivotCache>()) {
                if (cache.CacheId == null) continue;
                var relId = cache.Id?.Value;
                if (relId == null) continue;
                if (relId.Length == 0) continue;
                if (workbookPart.GetPartById(relId) is PivotTableCacheDefinitionPart part && part.PivotCacheDefinition != null) {
                    map[cache.CacheId.Value] = part.PivotCacheDefinition;
                }
            }
            return map;
        }

        private static List<string> BuildCacheFieldNames(PivotCacheDefinition? cacheDef) {
            var names = new List<string>();
            if (cacheDef?.CacheFields == null) return names;
            var used = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            int i = 0;
            foreach (var field in cacheDef.CacheFields.Elements<CacheField>()) {
                var name = field.Name?.Value ?? $"Field{i + 1}";
                if (string.IsNullOrWhiteSpace(name)) name = $"Field{i + 1}";
                name = EnsureUniqueName(name, used);
                used.Add(name);
                names.Add(name);
                i++;
            }
            return names;
        }

        private static Dictionary<uint, string> BuildNumberFormatCodeMap(WorkbookPart workbookPart) {
            var map = new Dictionary<uint, string>();
            foreach (var builtInFormat in ExcelBuiltInNumberFormats.Codes) {
                map[builtInFormat.Key] = builtInFormat.Value;
            }

            var numberingFormats = workbookPart.WorkbookStylesPart?.Stylesheet?.NumberingFormats;
            if (numberingFormats == null) return map;

            foreach (var format in numberingFormats.Elements<NumberingFormat>()) {
                if (format.NumberFormatId?.Value is uint id && format.FormatCode?.Value is string code) {
                    map[id] = code;
                }
            }

            return map;
        }

        private static List<IReadOnlyList<string>> BuildCacheFieldItems(PivotCacheDefinition? cacheDef) {
            var fields = new List<IReadOnlyList<string>>();
            if (cacheDef?.CacheFields == null) return fields;

            foreach (var field in cacheDef.CacheFields.Elements<CacheField>()) {
                var values = new List<string>();
                SharedItems? sharedItems = field.SharedItems;
                if (sharedItems != null) {
                    foreach (OpenXmlElement item in sharedItems.ChildElements) {
                        string? text = item switch {
                            StringItem stringItem => stringItem.Val?.Value,
                            NumberItem numberItem => numberItem.Val?.Value.ToString(System.Globalization.CultureInfo.InvariantCulture),
                            DateTimeItem dateItem => dateItem.Val?.Value.ToString("O", System.Globalization.CultureInfo.InvariantCulture),
                            BooleanItem booleanItem => booleanItem.Val?.Value.ToString(),
                            MissingItem => string.Empty,
                            _ => item.InnerText
                        };
                        values.Add(text ?? string.Empty);
                    }
                }

                fields.Add(values);
            }

            return fields;
        }

        private static List<string> ResolveFieldNames(IEnumerable<Field>? fields, IReadOnlyList<string> cacheFields) {
            var list = new List<string>();
            if (fields == null) return list;
            foreach (var field in fields) {
                if (field.Index == null) continue;
                list.Add(ResolveFieldName(field.Index.Value, cacheFields));
            }
            return list;
        }

        private static List<string> ResolvePageFieldNames(IEnumerable<PageField>? fields, IReadOnlyList<string> cacheFields) {
            var list = new List<string>();
            if (fields == null) return list;
            foreach (var field in fields) {
                if (field.Field == null) continue;
                list.Add(ResolveFieldName(field.Field.Value, cacheFields));
            }
            return list;
        }

        private static Dictionary<int, string> ResolveSelectedPageItems(IEnumerable<PageField>? fields, IReadOnlyList<IReadOnlyList<string>> cacheFieldItems) {
            var map = new Dictionary<int, string>();
            if (fields == null) return map;

            foreach (var field in fields) {
                if (field.Field == null || field.Item == null) continue;
                int fieldIndex = field.Field.Value;
                int itemIndex = (int)field.Item.Value;
                if (fieldIndex < 0 || fieldIndex >= cacheFieldItems.Count) continue;

                var items = cacheFieldItems[fieldIndex];
                if (itemIndex >= 0 && itemIndex < items.Count) {
                    map[fieldIndex] = items[itemIndex];
                }
            }

            return map;
        }

        private static List<ExcelPivotFieldInfo> ResolveFieldInfos(IEnumerable<PivotField>? fields, IReadOnlyList<string> cacheFields,
            IReadOnlyList<IReadOnlyList<string>> cacheFieldItems,
            IReadOnlyDictionary<int, string> selectedPageItems,
            IReadOnlyDictionary<uint, string> numberFormatCodes) {
            var list = new List<ExcelPivotFieldInfo>();
            if (fields == null) return list;
            int index = 0;
            foreach (var field in fields) {
                IReadOnlyList<string> itemValues = index < cacheFieldItems.Count ? cacheFieldItems[index] : Array.Empty<string>();
                selectedPageItems.TryGetValue(index, out string? selectedItem);
                uint? numberFormatId = field.NumberFormatId?.Value;
                list.Add(new ExcelPivotFieldInfo(
                    fieldName: ResolveFieldName(index, cacheFields),
                    axis: field.Axis?.Value,
                    sortType: field.SortType?.Value,
                    numberFormatId: numberFormatId,
                    showAll: field.ShowAll?.Value,
                    defaultSubtotal: field.DefaultSubtotal?.Value,
                    subtotalTop: field.SubtotalTop?.Value,
                    insertBlankRow: field.InsertBlankRow?.Value,
                    insertPageBreak: field.InsertPageBreak?.Value,
                    compact: field.Compact?.Value,
                    outline: field.Outline?.Value,
                    showDropDowns: field.ShowDropDowns?.Value,
                    multipleItemSelectionAllowed: field.MultipleItemSelectionAllowed?.Value,
                    includeNewItemsInFilter: field.IncludeNewItemsInFilter?.Value,
                    subtotalCaption: field.SubtotalCaption?.Value,
                    hiddenItems: ResolveHiddenItems(field.Items, itemValues),
                    selectedItem: selectedItem,
                    visibleItems: ResolveVisibleItems(field.Items, itemValues),
                    numberFormatCode: ResolveNumberFormatCode(numberFormatId, numberFormatCodes)));
                index++;
            }

            return list;
        }

        private static IReadOnlyList<string> ResolveHiddenItems(Items? items, IReadOnlyList<string> values) {
            if (items == null || values.Count == 0) return Array.Empty<string>();
            var hidden = new List<string>();
            foreach (var item in items.Elements<Item>()) {
                if (item.Hidden?.Value != true || item.Index == null) continue;
                int idx = (int)item.Index.Value;
                if (idx >= 0 && idx < values.Count) {
                    hidden.Add(values[idx]);
                }
            }

            return hidden;
        }

        private static IReadOnlyList<string> ResolveVisibleItems(Items? items, IReadOnlyList<string> values) {
            if (items == null || values.Count == 0) return Array.Empty<string>();
            var visible = new List<string>();
            foreach (var item in items.Elements<Item>()) {
                if (item.Index == null || item.Hidden?.Value == true) continue;
                int idx = (int)item.Index.Value;
                if (idx >= 0 && idx < values.Count) {
                    visible.Add(values[idx]);
                }
            }

            return visible;
        }

        private static List<ExcelPivotDataFieldInfo> ResolveDataFields(IEnumerable<DataField>? fields, IReadOnlyList<string> cacheFields,
            IReadOnlyDictionary<uint, string> numberFormatCodes) {
            var list = new List<ExcelPivotDataFieldInfo>();
            if (fields == null) return list;
            foreach (var field in fields) {
                int idx = field.Field?.Value is uint u ? (int)u : 0;
                var name = ResolveFieldName(idx, cacheFields);
                var fn = field.Subtotal?.Value ?? DataConsolidateFunctionValues.Sum;
                var display = field.Name?.Value;
                uint? numberFormatId = field.NumberFormatId?.Value;
                list.Add(new ExcelPivotDataFieldInfo(name, fn, display, numberFormatId,
                    ResolveNumberFormatCode(numberFormatId, numberFormatCodes),
                    field.ShowDataAs?.Value,
                    field.BaseField?.Value,
                    field.BaseItem?.Value));
            }
            return list;
        }

        private static List<ExcelPivotFilterInfo> ResolvePivotFilterInfos(IEnumerable<PivotFilter>? filters,
            IReadOnlyList<string> cacheFields,
            IReadOnlyList<ExcelPivotDataFieldInfo> dataFields) {
            var list = new List<ExcelPivotFilterInfo>();
            if (filters == null) return list;

            foreach (var filter in filters) {
                int fieldIndex = filter.Field?.Value is uint field ? (int)field : -1;
                string fieldName = ResolveFieldName(fieldIndex, cacheFields);
                string? dataFieldName = null;
                if (filter.MeasureField?.Value is uint measureField && measureField < dataFields.Count) {
                    var dataField = dataFields[(int)measureField];
                    dataFieldName = dataField.DisplayName ?? dataField.FieldName;
                }

                var top10 = filter.AutoFilter?
                    .Elements<FilterColumn>()
                    .Select(column => column.GetFirstChild<Top10>())
                    .FirstOrDefault(element => element != null);

                list.Add(new ExcelPivotFilterInfo(
                    fieldName,
                    filter.Type?.Value,
                    filter.StringValue1?.Value ?? FormatOpenXmlDouble(top10?.Val?.Value),
                    filter.StringValue2?.Value,
                    dataFieldName,
                    filter.Name?.Value,
                    filter.Description?.Value,
                    top10?.Top?.Value,
                    top10?.Percent?.Value,
                    FormatOpenXmlDouble(top10?.FilterValue?.Value)));
            }

            return list;
        }

        private static List<ExcelPivotCalculatedFieldInfo> ResolveCalculatedFieldInfos(PivotCacheDefinition? cacheDef,
            IReadOnlyDictionary<uint, string> numberFormatCodes) {
            var list = new List<ExcelPivotCalculatedFieldInfo>();
            if (cacheDef?.CacheFields == null) return list;

            foreach (var field in cacheDef.CacheFields.Elements<CacheField>()) {
                string? formula = field.Formula?.Value;
                if (string.IsNullOrWhiteSpace(formula)) continue;
                uint? numberFormatId = field.NumberFormatId?.Value;

                list.Add(new ExcelPivotCalculatedFieldInfo(
                    field.Name?.Value ?? string.Empty,
                    formula!,
                    field.Caption?.Value,
                    numberFormatId,
                    ResolveNumberFormatCode(numberFormatId, numberFormatCodes)));
            }

            return list;
        }

        private static string? ResolveNumberFormatCode(uint? numberFormatId, IReadOnlyDictionary<uint, string> numberFormatCodes) {
            return numberFormatId.HasValue && numberFormatCodes.TryGetValue(numberFormatId.Value, out string? code) ? code : null;
        }

        private static List<ExcelPivotGroupingInfo> ResolvePivotGroupingInfos(PivotCacheDefinition? cacheDef, IReadOnlyList<string> cacheFields) {
            var list = new List<ExcelPivotGroupingInfo>();
            if (cacheDef?.CacheFields == null) return list;

            int index = 0;
            foreach (var field in cacheDef.CacheFields.Elements<CacheField>()) {
                RangeProperties? range = field.FieldGroup?.GetFirstChild<RangeProperties>();
                if (range != null) {
                    list.Add(new ExcelPivotGroupingInfo(
                        ResolveFieldName(index, cacheFields),
                        range.GroupBy?.Value,
                        range.StartDate?.Value,
                        range.EndDate?.Value,
                        range.StartNumber?.Value,
                        range.EndNum?.Value,
                        range.GroupInterval?.Value,
                        range.AutoStart?.Value,
                        range.AutoEnd?.Value,
                        ResolveGroupItems(field.FieldGroup?.GetFirstChild<GroupItems>()),
                        field.FieldGroup?.Base?.Value,
                        field.FieldGroup?.ParentId?.Value));
                }

                index++;
            }

            return list;
        }

        private static IReadOnlyList<string> ResolveGroupItems(GroupItems? groupItems) {
            if (groupItems == null) return Array.Empty<string>();

            var values = new List<string>();
            foreach (OpenXmlElement item in groupItems.ChildElements) {
                string? text = item switch {
                    StringItem stringItem => stringItem.Val?.Value,
                    NumberItem numberItem => numberItem.Val?.Value.ToString(CultureInfo.InvariantCulture),
                    DateTimeItem dateItem => dateItem.Val?.Value.ToString("O", CultureInfo.InvariantCulture),
                    BooleanItem booleanItem => booleanItem.Val?.Value.ToString(),
                    DateGroupItem dateGroupItem => FormatDateGroupItem(dateGroupItem),
                    MissingItem => string.Empty,
                    _ => item.InnerText
                };
                values.Add(text ?? string.Empty);
            }

            return values;
        }

        private static string FormatDateGroupItem(DateGroupItem item) {
            var grouping = item.DateTimeGrouping?.Value.ToString() ?? "Date";
            var parts = new List<string>();
            if (item.Year?.Value is ushort year) parts.Add(year.ToString(CultureInfo.InvariantCulture));
            if (item.Month?.Value is ushort month) parts.Add(month.ToString(CultureInfo.InvariantCulture));
            if (item.Day?.Value is ushort day) parts.Add(day.ToString(CultureInfo.InvariantCulture));
            if (item.Hour?.Value is ushort hour) parts.Add(hour.ToString(CultureInfo.InvariantCulture));
            if (item.Minute?.Value is ushort minute) parts.Add(minute.ToString(CultureInfo.InvariantCulture));
            if (item.Second?.Value is ushort second) parts.Add(second.ToString(CultureInfo.InvariantCulture));
            return parts.Count == 0 ? grouping : $"{grouping}:{string.Join("-", parts)}";
        }

        private static string? FormatOpenXmlDouble(double? value) {
            return value?.ToString("G17", CultureInfo.InvariantCulture);
        }

        private static string ResolveFieldName(int index, IReadOnlyList<string> cacheFields) {
            if (index >= 0 && index < cacheFields.Count) return cacheFields[index];
            return $"Field{index + 1}";
        }

        private int ResolveSheetIndex(WorkbookPart workbookPart) {
            var workbook = workbookPart.Workbook ?? throw new InvalidOperationException("Workbook is missing.");
            var sheets = workbook.Sheets?.OfType<Sheet>().ToList();
            if (sheets == null) return -1;
            for (int i = 0; i < sheets.Count; i++) {
                if (ReferenceEquals(sheets[i], _sheet)) return i;
            }
            return -1;
        }

        private static ExcelPivotLayout ResolveLayout(BooleanValue? compactData, BooleanValue? outlineData) {
            if (outlineData != null && outlineData.Value) return ExcelPivotLayout.Outline;
            if (compactData != null && compactData.Value) return ExcelPivotLayout.Compact;
            return ExcelPivotLayout.Tabular;
        }
    }
}
