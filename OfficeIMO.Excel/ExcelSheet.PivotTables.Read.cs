using System.Globalization;
using System.Xml;
using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.Excel.OpenXml;
using OfficeIMO.OpenXml.Internal;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        private const string PivotCacheDefinitionRelationship = OfficeXmlNamespaces.TransitionalRelationshipTypePrefix + "pivotCacheDefinition";

        /// <summary>Returns pivot tables defined on this worksheet.</summary>
        public IReadOnlyList<ExcelPivotTableInfo> GetPivotTables() =>
            _excelDocument.ExecuteReadWithoutMaterializing(() => {
                if (DeferredEntry is { } entry && _excelDocument.TryReadWorksheetPackage(entry, ReadOwnedPivotTables,
                    out IReadOnlyList<ExcelPivotTableInfo> owned)) return owned;

                ExcelPartRef workbook = ExcelPartRef.Workbook(_excelDocument);
                return ReadPivotTables(workbook.Root, workbook.WorkbookStyles.Root,
                    ExcelPartRef.WorksheetForMetadata(this).PartsOfKind(ExcelPartKind.Pivot).Select(part => part.Root),
                    id => workbook.GetPartById(id).Root, Name);
            });

        private static IReadOnlyList<ExcelPivotTableInfo> ReadOwnedPivotTables(OfficeOpcPackage package, ExcelDeferredSheet sheet) {
            OfficeOpcRelationship[] pivots = package.GetRelationships(sheet.PartName)
                .Where(relationship => relationship.Type == PivotTableRelationship).ToArray();
            if (pivots.Length == 0) return Array.Empty<ExcelPivotTableInfo>();

            var buffers = new List<OfficePooledBytes>();
            OfficeXmlNode ReadRoot(string? target) {
                if (target == null || !package.PartExists(target)) return default;
                OfficePooledBytes bytes = package.ReadPart(target);
                buffers.Add(bytes);
                OfficeXmlTree tree = OfficeXmlTree.Parse(bytes.Array, bytes.Length, package.Namespaces);
                return new OfficeXmlNode(tree, tree.Root);
            }

            try {
                string? workbook = package.GetRelatedPartName(null, OfficeOpcPackage.OfficeDocumentRelationshipType);
                if (workbook == null) throw new InvalidOperationException("Workbook is missing.");
                return ReadPivotTables(ReadRoot(workbook), ReadRoot(package.GetRelatedPartName(workbook,
                        OfficeXmlNamespaces.TransitionalRelationshipTypePrefix + "styles")),
                    pivots.Select(relationship => ReadRoot(relationship.TargetPartName)),
                    id => {
                        OfficeOpcRelationship? relationship = package.FindRelationship(workbook, id);
                        return relationship?.Type == PivotCacheDefinitionRelationship ? ReadRoot(relationship.TargetPartName) : default;
                    }, sheet.Name);
            } finally {
                foreach (OfficePooledBytes buffer in buffers) buffer.Dispose();
            }
        }

        private static IReadOnlyList<ExcelPivotTableInfo> ReadPivotTables(OfficeXmlNode workbook, OfficeXmlNode styles,
            IEnumerable<OfficeXmlNode> definitions, Func<string, OfficeXmlNode> readCache, string sheetName) {
            if (!workbook.Is(WorksheetMarkup, "workbook")) throw new InvalidOperationException("Workbook is missing.");
            var list = new List<ExcelPivotTableInfo>();
            Dictionary<uint, OfficeXmlNode>? cacheMap = null;
            Dictionary<uint, string>? numberFormatCodes = null;
            int sheetIndex = -1;
            int position = 0;
            foreach (OfficeXmlNode sheet in workbook.Child(WorksheetMarkup, "sheets").Children(WorksheetMarkup, "sheet")) {
                if (string.Equals(sheet.Attribute("name"), sheetName, StringComparison.Ordinal)) {
                    sheetIndex = position;
                    break;
                }
                position++;
            }

            foreach (OfficeXmlNode def in definitions) {
                if (!def.Is(WorksheetMarkup, "pivotTableDefinition")) continue;
                cacheMap ??= BuildPivotCacheMap(workbook, readCache);
                numberFormatCodes ??= BuildNumberFormatCodeMap(styles);
                uint cacheId = def.UInt32Attribute("cacheId") ?? 0U;
                cacheMap.TryGetValue(cacheId, out var cacheDef);
                var cacheFields = BuildCacheFieldNames(cacheDef);
                OfficeXmlNode worksheetSource = cacheDef.Child(WorksheetMarkup, "cacheSource").Child(WorksheetMarkup, "worksheetSource");
                var rowAxis = def.Child(WorksheetMarkup, "rowFields").Children(WorksheetMarkup, "field");
                var columnAxis = def.Child(WorksheetMarkup, "colFields").Children(WorksheetMarkup, "field");
                var pageAxis = def.Child(WorksheetMarkup, "pageFields").Children(WorksheetMarkup, "pageField");
                var sourceSheet = worksheetSource.Attribute("sheet");
                var sourceRange = worksheetSource.Attribute("ref");

                var rowFields = ResolveFieldNames(rowAxis, cacheFields);
                var columnFields = ResolveFieldNames(columnAxis, cacheFields);
                var pageFields = ResolvePageFieldNames(pageAxis, cacheFields);
                var dataFields = ResolveDataFields(def.Child(WorksheetMarkup, "dataFields").Children(WorksheetMarkup, "dataField"), cacheFields, numberFormatCodes);
                var cacheFieldItems = BuildCacheFieldItems(cacheDef);
                var selectedPageItems = ResolveSelectedPageItems(pageAxis, cacheFieldItems);
                var fieldInfos = ResolveFieldInfos(def.Child(WorksheetMarkup, "pivotFields").Children(WorksheetMarkup, "pivotField"), cacheFields, cacheFieldItems, selectedPageItems, numberFormatCodes);
                var filterInfos = ResolvePivotFilterInfos(def.Child(WorksheetMarkup, "filters").Children(WorksheetMarkup, "filter"), cacheFields, dataFields);
                var calculatedFieldInfos = ResolveCalculatedFieldInfos(cacheDef, numberFormatCodes);
                var groupingInfos = ResolvePivotGroupingInfos(cacheDef, cacheFields);
                var rowAxisIndices = ResolvePivotAxisIndices(rowAxis);
                var columnAxisIndices = ResolvePivotAxisIndices(columnAxis);
                int valuesFieldCount = rowAxisIndices.Count(f => f == -2) + columnAxisIndices.Count(f => f == -2);
                ExcelPivotTableAxis? valuesAxis = valuesFieldCount == 1
                    ? rowAxisIndices.Contains(-2) ? ExcelPivotTableAxis.AxisRow : ExcelPivotTableAxis.AxisColumn : null;

                var layout = ResolveLayout(def.BooleanAttribute("compactData"), def.BooleanAttribute("outlineData"));

                list.Add(new ExcelPivotTableInfo(
                    name: def.Attribute("name") ?? string.Empty,
                    cacheId: cacheId,
                    location: def.Child(WorksheetMarkup, "location").Attribute("ref"),
                    sourceSheet: sourceSheet,
                    sourceRange: sourceRange,
                    sheetName: sheetName,
                    sheetIndex: sheetIndex,
                    pivotStyle: def.Attribute("pivotTableStyle"),
                    layout: layout,
                    dataOnRows: def.BooleanAttribute("dataOnRows"),
                    showHeaders: def.BooleanAttribute("showHeaders"),
                    showEmptyRows: def.BooleanAttribute("showEmptyRow"),
                    showEmptyColumns: def.BooleanAttribute("showEmptyCol"),
                    showDrill: def.BooleanAttribute("showDrill"),
                    rowGrandTotals: def.BooleanAttribute("rowGrandTotals"),
                    columnGrandTotals: def.BooleanAttribute("colGrandTotals"),
                    rowHeaderCaption: def.Attribute("rowHeaderCaption"),
                    columnHeaderCaption: def.Attribute("colHeaderCaption"),
                    grandTotalCaption: def.Attribute("grandTotalCaption"),
                    missingCaption: def.Attribute("missingCaption"),
                    errorCaption: def.Attribute("errorCaption"),
                    showDataDropDown: def.BooleanAttribute("showDataDropDown"),
                    showDropZones: def.BooleanAttribute("showDropZones"),
                    showDataTips: def.BooleanAttribute("showDataTips"),
                    showMemberPropertyTips: def.BooleanAttribute("showMemberPropertyTips"),
                    fieldListSortAscending: def.BooleanAttribute("fieldListSortAscending"),
                    customListSort: def.BooleanAttribute("customListSort"),
                    rowFields: rowFields,
                    columnFields: columnFields,
                    pageFields: pageFields,
                    dataFields: dataFields,
                    fields: fieldInfos,
                    filters: filterInfos,
                    calculatedFields: calculatedFieldInfos,
                    groupings: groupingInfos,
                    refreshOnOpen: cacheDef.BooleanAttribute("refreshOnLoad"),
                    saveSourceData: cacheDef.BooleanAttribute("saveData"),
                    preserveFormatting: def.BooleanAttribute("preserveFormatting"),
                    enableDrill: def.BooleanAttribute("enableDrill"),
                    hasValuesAxisField: valuesFieldCount > 0) {
                    RowSourceFields = ResolveFieldNames(rowAxis, cacheFields, sourceFieldsOnly: true),
                    ColumnSourceFields = ResolveFieldNames(columnAxis, cacheFields, sourceFieldsOnly: true),
                    ValuesAxis = valuesAxis,
                    ValuesAxisPosition = valuesAxis.HasValue ? Array.IndexOf(valuesAxis == ExcelPivotTableAxis.AxisRow ? rowAxisIndices : columnAxisIndices, -2) : null
                });
            }
            return list;
        }

        private static Dictionary<uint, OfficeXmlNode> BuildPivotCacheMap(OfficeXmlNode workbook, Func<string, OfficeXmlNode> readCache) {
            var map = new Dictionary<uint, OfficeXmlNode>();
            foreach (OfficeXmlNode cache in workbook.Child(WorksheetMarkup, "pivotCaches").Children(WorksheetMarkup, "pivotCache")) {
                string? id = cache.Attribute(OfficeXmlNamespaces.Relationships, "id");
                if (string.IsNullOrEmpty(id) || cache.UInt32Attribute("cacheId") is not uint cacheId) continue;
                OfficeXmlNode definition = readCache(id!);
                if (definition.Is(WorksheetMarkup, "pivotCacheDefinition")) map[cacheId] = definition;
            }
            return map;
        }

        private static OfficeXmlNode.ChildList ReadCacheFields(OfficeXmlNode definition) =>
            definition.Child(WorksheetMarkup, "cacheFields").Children(WorksheetMarkup, "cacheField");

        private static List<string> BuildCacheFieldNames(OfficeXmlNode cacheDef) {
            var names = new List<string>();
            var used = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (OfficeXmlNode field in ReadCacheFields(cacheDef)) {
                string? name = field.Attribute("name");
                if (string.IsNullOrWhiteSpace(name)) name = $"Field{names.Count + 1}";
                name = EnsureUniqueName(name!, used);
                used.Add(name);
                names.Add(name);
            }
            return names;
        }

        private static Dictionary<uint, string> BuildNumberFormatCodeMap(OfficeXmlNode styles) {
            var map = new Dictionary<uint, string>();
            foreach (var builtInFormat in ExcelBuiltInNumberFormats.Codes) map[builtInFormat.Key] = builtInFormat.Value;
            foreach (OfficeXmlNode format in styles.Child(WorksheetMarkup, "numFmts").Children(WorksheetMarkup, "numFmt")) {
                if (format.UInt32Attribute("numFmtId") is uint id && format.Attribute("formatCode") is string code) map[id] = code;
            }
            return map;
        }

        private static List<IReadOnlyList<string>> BuildCacheFieldItems(OfficeXmlNode cacheDef) {
            var fields = new List<IReadOnlyList<string>>();
            foreach (OfficeXmlNode field in ReadCacheFields(cacheDef)) {
                OfficeXmlNode items = GroupItems(field);
                if (items.IsNone) items = field.Child(WorksheetMarkup, "sharedItems");
                fields.Add(ResolveGroupItems(items, includeDateGroup: false));
            }
            return fields;
        }

        private static int[] ResolvePivotAxisIndices(OfficeXmlNode.ChildList fields) {
            var indices = new List<int>();
            foreach (OfficeXmlNode field in fields) indices.Add(field.Int32Attribute("x") ?? int.MinValue);
            return indices.ToArray();
        }

        private static List<string> ResolveFieldNames(OfficeXmlNode.ChildList fields, IReadOnlyList<string> cacheFields, bool sourceFieldsOnly = false) {
            var list = new List<string>();
            foreach (OfficeXmlNode field in fields) {
                if (field.Int32Attribute("x") is int index && (!sourceFieldsOnly || index >= 0)) list.Add(ResolveFieldName(index, cacheFields));
            }
            return list;
        }

        private static List<string> ResolvePageFieldNames(OfficeXmlNode.ChildList fields, IReadOnlyList<string> cacheFields) {
            var list = new List<string>();
            foreach (OfficeXmlNode field in fields) {
                if (field.Int32Attribute("fld") is int index) list.Add(ResolveFieldName(index, cacheFields));
            }
            return list;
        }

        private static Dictionary<int, string> ResolveSelectedPageItems(OfficeXmlNode.ChildList fields, IReadOnlyList<IReadOnlyList<string>> cacheFieldItems) {
            var map = new Dictionary<int, string>();
            foreach (OfficeXmlNode field in fields) {
                if (field.Int32Attribute("fld") is int index && index >= 0 && index < cacheFieldItems.Count
                    && field.UInt32Attribute("item") is uint item && item < cacheFieldItems[index].Count) {
                    map[index] = cacheFieldItems[index][(int)item];
                }
            }
            return map;
        }

        private static List<ExcelPivotFieldInfo> ResolveFieldInfos(OfficeXmlNode.ChildList fields, IReadOnlyList<string> cacheFields,
            IReadOnlyList<IReadOnlyList<string>> cacheFieldItems, IReadOnlyDictionary<int, string> selectedPageItems,
            IReadOnlyDictionary<uint, string> numberFormatCodes) {
            var list = new List<ExcelPivotFieldInfo>();
            foreach (OfficeXmlNode field in fields) {
                int index = list.Count;
                IReadOnlyList<string> itemValues = index < cacheFieldItems.Count ? cacheFieldItems[index] : Array.Empty<string>();
                selectedPageItems.TryGetValue(index, out string? selectedItem);
                uint? numberFormatId = field.UInt32Attribute("numFmtId");
                OfficeXmlNode items = field.Child(WorksheetMarkup, "items");
                list.Add(new ExcelPivotFieldInfo(
                    fieldName: ResolveFieldName(index, cacheFields),
                    axis: PivotValue<ExcelPivotTableAxis>(field, "axis", ExcelOpenXmlValueTypeExtensions.TryParseOpenXmlValue),
                    sortType: PivotValue<ExcelPivotFieldSort>(field, "sortType", ExcelOpenXmlValueTypeExtensions.TryParseOpenXmlValue),
                    numberFormatId: numberFormatId,
                    showAll: field.BooleanAttribute("showAll"),
                    defaultSubtotal: field.BooleanAttribute("defaultSubtotal"),
                    subtotalTop: field.BooleanAttribute("subtotalTop"),
                    insertBlankRow: field.BooleanAttribute("insertBlankRow"),
                    insertPageBreak: field.BooleanAttribute("insertPageBreak"),
                    compact: field.BooleanAttribute("compact"),
                    outline: field.BooleanAttribute("outline"),
                    showDropDowns: field.BooleanAttribute("showDropDowns"),
                    multipleItemSelectionAllowed: field.BooleanAttribute("multipleItemSelectionAllowed"),
                    includeNewItemsInFilter: field.BooleanAttribute("includeNewItemsInFilter"),
                    subtotalCaption: field.Attribute("subtotalCaption"),
                    hiddenItems: ResolvePivotItems(items, itemValues, hidden: true),
                    selectedItem: selectedItem,
                    visibleItems: ResolvePivotItems(items, itemValues, hidden: false),
                    numberFormatCode: ResolveNumberFormatCode(numberFormatId, numberFormatCodes)));
            }
            return list;
        }

        private static IReadOnlyList<string> ResolvePivotItems(OfficeXmlNode items, IReadOnlyList<string> values, bool hidden) {
            if (items.IsNone || values.Count == 0) return Array.Empty<string>();
            var result = new List<string>();
            foreach (OfficeXmlNode item in items.Children(WorksheetMarkup, "item")) {
                if ((item.BooleanAttribute("h") == true) == hidden && item.UInt32Attribute("x") is uint index && index < values.Count) {
                    result.Add(values[(int)index]);
                }
            }
            return result;
        }

        private static List<ExcelPivotDataFieldInfo> ResolveDataFields(OfficeXmlNode.ChildList fields, IReadOnlyList<string> cacheFields,
            IReadOnlyDictionary<uint, string> numberFormatCodes) {
            var list = new List<ExcelPivotDataFieldInfo>();
            foreach (OfficeXmlNode field in fields) {
                int index = unchecked((int)(field.UInt32Attribute("fld") ?? 0U));
                uint? format = field.UInt32Attribute("numFmtId");
                list.Add(new ExcelPivotDataFieldInfo(ResolveFieldName(index, cacheFields),
                    PivotValue<ExcelPivotDataFunction>(field, "subtotal", ExcelOpenXmlValueTypeExtensions.TryParseOpenXmlValue) ?? ExcelPivotDataFunction.Sum,
                    field.Attribute("name"), format, ResolveNumberFormatCode(format, numberFormatCodes),
                    PivotValue<ExcelPivotShowDataAs>(field, "showDataAs", ExcelOpenXmlValueTypeExtensions.TryParseOpenXmlValue),
                    field.Int32Attribute("baseField"), field.UInt32Attribute("baseItem")));
            }
            return list;
        }

        private static List<ExcelPivotFilterInfo> ResolvePivotFilterInfos(OfficeXmlNode.ChildList filters,
            IReadOnlyList<string> cacheFields, IReadOnlyList<ExcelPivotDataFieldInfo> dataFields) {
            var list = new List<ExcelPivotFilterInfo>();
            foreach (OfficeXmlNode filter in filters) {
                int index = unchecked((int)(filter.UInt32Attribute("fld") ?? uint.MaxValue));
                string? dataFieldName = null;
                if (filter.UInt32Attribute("iMeasureFld") is uint measure && measure < dataFields.Count) {
                    var data = dataFields[(int)measure];
                    dataFieldName = data.DisplayName ?? data.FieldName;
                }

                OfficeXmlNode top10 = default;
                foreach (OfficeXmlNode column in filter.Child(WorksheetMarkup, "autoFilter").Children(WorksheetMarkup, "filterColumn")) {
                    top10 = column.Child(WorksheetMarkup, "top10");
                    if (top10.Exists) break;
                }
                ExcelPivotFilterType? type = PivotValue<ExcelPivotFilterType>(filter, "type", ExcelOpenXmlValueTypeExtensions.TryParseOpenXmlValue);
                list.Add(new ExcelPivotFilterInfo(ResolveFieldName(index, cacheFields), type,
                    filter.Attribute("stringValue1") ?? FormatOpenXmlDouble(top10.DoubleAttribute("val")),
                    filter.Attribute("stringValue2"), dataFieldName, filter.Attribute("name"), filter.Attribute("description"),
                    top10.BooleanAttribute("top"), top10.BooleanAttribute("percent"), FormatOpenXmlDouble(top10.DoubleAttribute("filterVal")),
                    IsMaterializedPivotFixedDateFilter(type) && QualifiedPivotWholeDay(filter)));
            }
            return list;
        }

        private static List<ExcelPivotCalculatedFieldInfo> ResolveCalculatedFieldInfos(OfficeXmlNode cacheDef,
            IReadOnlyDictionary<uint, string> numberFormatCodes) {
            var list = new List<ExcelPivotCalculatedFieldInfo>();
            foreach (OfficeXmlNode field in ReadCacheFields(cacheDef)) {
                string? formula = field.Attribute("formula");
                if (string.IsNullOrWhiteSpace(formula)) continue;
                uint? format = field.UInt32Attribute("numFmtId");
                list.Add(new ExcelPivotCalculatedFieldInfo(field.Attribute("name") ?? string.Empty, formula!, field.Attribute("caption"),
                    format, ResolveNumberFormatCode(format, numberFormatCodes)));
            }
            return list;
        }

        private static string? ResolveNumberFormatCode(uint? format, IReadOnlyDictionary<uint, string> numberFormatCodes) =>
            format.HasValue && numberFormatCodes.TryGetValue(format.Value, out string? code) ? code : null;

        private static List<ExcelPivotGroupingInfo> ResolvePivotGroupingInfos(OfficeXmlNode cacheDef, IReadOnlyList<string> cacheFields) {
            var list = new List<ExcelPivotGroupingInfo>();
            int index = 0;
            foreach (OfficeXmlNode field in ReadCacheFields(cacheDef)) {
                OfficeXmlNode group = field.Child(WorksheetMarkup, "fieldGroup");
                OfficeXmlNode range = group.Child(WorksheetMarkup, "rangePr");
                if (range.Exists || group.Child(WorksheetMarkup, "discretePr").Exists) {
                    list.Add(new ExcelPivotGroupingInfo(ResolveFieldName(index, cacheFields),
                        PivotValue<ExcelPivotGroupBy>(range, "groupBy", ExcelOpenXmlValueTypeExtensions.TryParseOpenXmlValue),
                        ReadPivotDate(range.Attribute("startDate")), ReadPivotDate(range.Attribute("endDate")),
                        range.DoubleAttribute("startNum"), range.DoubleAttribute("endNum"), range.DoubleAttribute("groupInterval"),
                        range.BooleanAttribute("autoStart"), range.BooleanAttribute("autoEnd"), ResolveGroupItems(GroupItems(field), includeDateGroup: true),
                        group.UInt32Attribute("base"), group.UInt32Attribute("par")));
                }
                index++;
            }
            return list;
        }

        private static OfficeXmlNode GroupItems(OfficeXmlNode field) => field.Child(WorksheetMarkup, "fieldGroup").Child(WorksheetMarkup, "groupItems");

        private static IReadOnlyList<string> ResolveGroupItems(OfficeXmlNode items, bool includeDateGroup) {
            if (items.IsNone) return Array.Empty<string>();
            var values = new List<string>();
            foreach (OfficeXmlNode item in items.Children()) {
                if (item.IsExtensionMarkup()) continue;
                string? value = item.NamespaceId == WorksheetMarkup ? item.LocalName switch {
                    "s" => item.Attribute("v"),
                    "n" => item.DoubleAttribute("v")?.ToString(CultureInfo.InvariantCulture),
                    "d" => ReadPivotDate(item.Attribute("v"))?.ToString("O", CultureInfo.InvariantCulture),
                    "b" => item.BooleanAttribute("v")?.ToString(),
                    "dateGroupItem" when includeDateGroup => FormatDateGroupItem(item),
                    "m" => string.Empty,
                    _ => item.InnerText
                } : item.InnerText;
                values.Add(value ?? string.Empty);
            }
            return values;
        }

        private static string FormatDateGroupItem(OfficeXmlNode item) {
            string grouping = item.Attribute("dateTimeGrouping") ?? "Date";
            var parts = new List<string>();
            foreach (string attribute in new[] { "year", "month", "day", "hour", "minute", "second" }) {
                if (item.UInt32Attribute(attribute) is uint number && number <= ushort.MaxValue) parts.Add(number.ToString(CultureInfo.InvariantCulture));
            }
            return parts.Count == 0 ? grouping : $"{grouping}:{string.Join("-", parts)}";
        }

        private static DateTime? ReadPivotDate(string? value) =>
            value == null ? (DateTime?)null : XmlConvert.ToDateTime(value, XmlDateTimeSerializationMode.RoundtripKind);

        private delegate bool TryParsePivotValue<T>(string token, out T result) where T : struct;
        private static T? PivotValue<T>(OfficeXmlNode element, string attribute, TryParsePivotValue<T> parse) where T : struct {
            string? token = element.Attribute(attribute);
            if (token == null) return null;
            if (parse(token, out T result)) return result;
            throw new FormatException($"Invalid pivot {attribute} value '{token}'.");
        }

        private static string? FormatOpenXmlDouble(double? value) => value.HasValue ? InvariantNumberText.Get(value.Value) : null;

        private static string ResolveFieldName(int index, IReadOnlyList<string> cacheFields) =>
            index >= 0 && index < cacheFields.Count ? cacheFields[index] : $"Field{index + 1}";

        private static ExcelPivotLayout ResolveLayout(bool? compactData, bool? outlineData) =>
            outlineData == true ? ExcelPivotLayout.Outline : compactData == true ? ExcelPivotLayout.Compact : ExcelPivotLayout.Tabular;
    }
}
