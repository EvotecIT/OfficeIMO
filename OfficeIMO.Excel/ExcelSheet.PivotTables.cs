using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Globalization;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        private static readonly IReadOnlyDictionary<uint, string> BuiltInNumberFormatCodes = new Dictionary<uint, string> {
            [0U] = "General",
            [1U] = "0",
            [2U] = "0.00",
            [3U] = "#,##0",
            [4U] = "#,##0.00",
            [9U] = "0%",
            [10U] = "0.00%",
            [11U] = "0.00E+00",
            [12U] = "# ?/?",
            [13U] = "# ??/??",
            [14U] = "mm-dd-yy",
            [15U] = "d-mmm-yy",
            [16U] = "d-mmm",
            [17U] = "mmm-yy",
            [18U] = "h:mm AM/PM",
            [19U] = "h:mm:ss AM/PM",
            [20U] = "h:mm",
            [21U] = "h:mm:ss",
            [22U] = "m/d/yy h:mm",
            [37U] = "#,##0 ;(#,##0)",
            [38U] = "#,##0 ;[Red](#,##0)",
            [39U] = "#,##0.00;(#,##0.00)",
            [40U] = "#,##0.00;[Red](#,##0.00)",
            [45U] = "mm:ss",
            [46U] = "[h]:mm:ss",
            [47U] = "mmss.0",
            [48U] = "##0.0E+0",
            [49U] = "@"
        };

        /// <summary>
        /// Returns pivot tables defined on this worksheet.
        /// </summary>
        public IReadOnlyList<ExcelPivotTableInfo> GetPivotTables() {
            return Locking.ExecuteRead(_excelDocument.EnsureLock(), () => {
                var list = new List<ExcelPivotTableInfo>();
                var workbookPart = WorkbookPartRoot;

                var cacheMap = BuildPivotCacheMap(workbookPart);
                var numberFormatCodes = BuildNumberFormatCodeMap(workbookPart);
                var sheetIndex = ResolveSheetIndex(workbookPart);

                foreach (var pivotPart in _worksheetPart.PivotTableParts) {
                    var def = pivotPart.PivotTableDefinition;
                    if (def == null) continue;

                    uint cacheId = def.CacheId?.Value ?? 0U;
                    cacheMap.TryGetValue(cacheId, out var cacheDef);
                    var cacheFields = BuildCacheFieldNames(cacheDef);
                    var sourceSheet = cacheDef?.CacheSource?.WorksheetSource?.Sheet?.Value;
                    var sourceRange = cacheDef?.CacheSource?.WorksheetSource?.Reference?.Value;

                    var rowFields = ResolveFieldNames(def.RowFields?.Elements<Field>(), cacheFields);
                    var columnFields = ResolveFieldNames(def.ColumnFields?.Elements<Field>(), cacheFields);
                    var pageFields = ResolvePageFieldNames(def.PageFields?.Elements<PageField>(), cacheFields);
                    var dataFields = ResolveDataFields(def.DataFields?.Elements<DataField>(), cacheFields, numberFormatCodes);
                    var cacheFieldItems = BuildCacheFieldItems(cacheDef);
                    var selectedPageItems = ResolveSelectedPageItems(def.PageFields?.Elements<PageField>(), cacheFieldItems);
                    var fieldInfos = ResolveFieldInfos(def.PivotFields?.Elements<PivotField>(), cacheFields, cacheFieldItems, selectedPageItems, numberFormatCodes);
                    var filterInfos = ResolvePivotFilterInfos(def.PivotFilters?.Elements<PivotFilter>(), cacheFields, dataFields);
                    var calculatedFieldInfos = ResolveCalculatedFieldInfos(cacheDef, numberFormatCodes);
                    var groupingInfos = ResolvePivotGroupingInfos(cacheDef, cacheFields);

                    var layout = ResolveLayout(def.CompactData, def.OutlineData);

                    list.Add(new ExcelPivotTableInfo(
                        name: def.Name?.Value ?? string.Empty,
                        cacheId: cacheId,
                        location: def.Location?.Reference?.Value,
                        sourceSheet: sourceSheet,
                        sourceRange: sourceRange,
                        sheetName: Name,
                        sheetIndex: sheetIndex,
                        pivotStyle: def.PivotTableStyleName?.Value,
                        layout: layout,
                        dataOnRows: def.DataOnRows?.Value,
                        showHeaders: def.ShowHeaders?.Value,
                        showEmptyRows: def.ShowEmptyRow?.Value,
                        showEmptyColumns: def.ShowEmptyColumn?.Value,
                        showDrill: def.ShowDrill?.Value,
                        rowGrandTotals: def.RowGrandTotals?.Value,
                        columnGrandTotals: def.ColumnGrandTotals?.Value,
                        rowHeaderCaption: def.RowHeaderCaption?.Value,
                        columnHeaderCaption: def.ColumnHeaderCaption?.Value,
                        grandTotalCaption: def.GrandTotalCaption?.Value,
                        missingCaption: def.MissingCaption?.Value,
                        errorCaption: def.ErrorCaption?.Value,
                        showDataDropDown: def.ShowDataDropDown?.Value,
                        showDropZones: def.ShowDropZones?.Value,
                        showDataTips: def.ShowDataTips?.Value,
                        showMemberPropertyTips: def.ShowMemberPropertyTips?.Value,
                        fieldListSortAscending: def.FieldListSortAscending?.Value,
                        customListSort: def.CustomListSort?.Value,
                        rowFields: rowFields,
                        columnFields: columnFields,
                        pageFields: pageFields,
                        dataFields: dataFields,
                        fields: fieldInfos,
                        filters: filterInfos,
                        calculatedFields: calculatedFieldInfos,
                        groupings: groupingInfos));
                }

                return list;
            });
        }

        /// <summary>
        /// Adds a basic pivot table based on a source range and places it at a destination cell.
        /// </summary>
        /// <param name="sourceRange">Source data range (including header row), e.g. "A1:D100".</param>
        /// <param name="destinationCell">Top-left cell for the pivot table (e.g. "F2").</param>
        /// <param name="name">Optional pivot table name. Defaults to "PivotTable1" style.</param>
        /// <param name="rowFields">Optional row fields (header names).</param>
        /// <param name="columnFields">Optional column fields (header names).</param>
        /// <param name="pageFields">Optional page fields (header names) used as filters.</param>
        /// <param name="dataFields">Optional data field definitions. Defaults to last column with Sum.</param>
        /// <param name="showRowGrandTotals">Show row grand totals.</param>
        /// <param name="showColumnGrandTotals">Show column grand totals.</param>
        /// <param name="pivotStyleName">Optional pivot table style name.</param>
        /// <param name="layout">Layout mode (Compact, Outline, Tabular).</param>
        /// <param name="dataOnRows">Whether to show data fields on rows instead of columns.</param>
        /// <param name="showHeaders">Whether to show field headers.</param>
        /// <param name="showEmptyRows">Whether to show empty rows.</param>
        /// <param name="showEmptyColumns">Whether to show empty columns.</param>
        /// <param name="showDrill">Whether to show drill indicators.</param>
        public void AddPivotTable(
            string sourceRange,
            string destinationCell,
            string? name,
            IEnumerable<string>? rowFields,
            IEnumerable<string>? columnFields,
            IEnumerable<string>? pageFields,
            IEnumerable<ExcelPivotDataField>? dataFields,
            bool showRowGrandTotals,
            bool showColumnGrandTotals,
            string? pivotStyleName,
            ExcelPivotLayout layout,
            bool? dataOnRows,
            bool? showHeaders,
            bool? showEmptyRows,
            bool? showEmptyColumns,
            bool? showDrill) {
            AddPivotTable(
                sourceRange,
                destinationCell,
                name,
                rowFields,
                columnFields,
                pageFields,
                dataFields,
                showRowGrandTotals,
                showColumnGrandTotals,
                pivotStyleName,
                layout,
                dataOnRows,
                showHeaders,
                showEmptyRows,
                showEmptyColumns,
                showDrill,
                fieldOptions: null,
                rowHeaderCaption: null,
                columnHeaderCaption: null,
                grandTotalCaption: null,
                missingCaption: null,
                errorCaption: null,
                showDataDropDown: null,
                showDropZones: null,
                showDataTips: null,
                showMemberPropertyTips: null,
                fieldListSortAscending: null,
                customListSort: null,
                pivotFilters: null,
                calculatedFields: null,
                groupings: null);
        }

        /// <summary>
        /// Adds a basic pivot table based on a source range and places it at a destination cell.
        /// </summary>
        /// <param name="sourceRange">Source data range (including header row), e.g. "A1:D100".</param>
        /// <param name="destinationCell">Top-left cell for the pivot table (e.g. "F2").</param>
        /// <param name="name">Optional pivot table name. Defaults to "PivotTable1" style.</param>
        /// <param name="rowFields">Optional row fields (header names).</param>
        /// <param name="columnFields">Optional column fields (header names).</param>
        /// <param name="pageFields">Optional page fields (header names) used as filters.</param>
        /// <param name="dataFields">Optional data field definitions. Defaults to last column with Sum.</param>
        /// <param name="showRowGrandTotals">Show row grand totals.</param>
        /// <param name="showColumnGrandTotals">Show column grand totals.</param>
        /// <param name="pivotStyleName">Optional pivot table style name.</param>
        /// <param name="layout">Layout mode (Compact, Outline, Tabular).</param>
        /// <param name="dataOnRows">Whether to show data fields on rows instead of columns.</param>
        /// <param name="showHeaders">Whether to show field headers.</param>
        /// <param name="showEmptyRows">Whether to show empty rows.</param>
        /// <param name="showEmptyColumns">Whether to show empty columns.</param>
        /// <param name="showDrill">Whether to show drill indicators.</param>
        /// <param name="fieldOptions">Optional formatting and display options for source fields.</param>
        /// <param name="rowHeaderCaption">Optional row header caption.</param>
        /// <param name="columnHeaderCaption">Optional column header caption.</param>
        /// <param name="grandTotalCaption">Optional grand total caption.</param>
        /// <param name="missingCaption">Optional caption for missing values.</param>
        /// <param name="errorCaption">Optional caption for error values.</param>
        /// <param name="showDataDropDown">Whether to show the data drop-down.</param>
        /// <param name="showDropZones">Whether to show drop zones.</param>
        /// <param name="showDataTips">Whether to show data tips.</param>
        /// <param name="showMemberPropertyTips">Whether to show member property tips.</param>
        /// <param name="fieldListSortAscending">Whether field list sorting is ascending.</param>
        /// <param name="customListSort">Whether custom-list sorting is enabled.</param>
        /// <param name="pivotFilters">Optional label and value filters.</param>
        /// <param name="calculatedFields">Optional formula-backed pivot cache fields.</param>
        /// <param name="groupings">Optional date or numeric grouping metadata.</param>
        public void AddPivotTable(
            string sourceRange,
            string destinationCell,
            string? name = null,
            IEnumerable<string>? rowFields = null,
            IEnumerable<string>? columnFields = null,
            IEnumerable<string>? pageFields = null,
            IEnumerable<ExcelPivotDataField>? dataFields = null,
            bool showRowGrandTotals = true,
            bool showColumnGrandTotals = true,
            string? pivotStyleName = null,
            ExcelPivotLayout layout = ExcelPivotLayout.Compact,
            bool? dataOnRows = null,
            bool? showHeaders = null,
            bool? showEmptyRows = null,
            bool? showEmptyColumns = null,
            bool? showDrill = null,
            IEnumerable<ExcelPivotFieldOptions>? fieldOptions = null,
            string? rowHeaderCaption = null,
            string? columnHeaderCaption = null,
            string? grandTotalCaption = null,
            string? missingCaption = null,
            string? errorCaption = null,
            bool? showDataDropDown = null,
            bool? showDropZones = null,
            bool? showDataTips = null,
            bool? showMemberPropertyTips = null,
            bool? fieldListSortAscending = null,
            bool? customListSort = null,
            IEnumerable<ExcelPivotFilter>? pivotFilters = null,
            IEnumerable<ExcelPivotCalculatedField>? calculatedFields = null,
            IEnumerable<ExcelPivotGrouping>? groupings = null) {
            if (string.IsNullOrWhiteSpace(sourceRange)) throw new ArgumentNullException(nameof(sourceRange));
            if (string.IsNullOrWhiteSpace(destinationCell)) throw new ArgumentNullException(nameof(destinationCell));
            if (!A1.TryParseRange(sourceRange, out int r1, out int c1, out int r2, out int c2)) {
                throw new ArgumentException($"Invalid A1 range '{sourceRange}'.", nameof(sourceRange));
            }

            var (destRow, destCol) = A1.ParseCellRef(destinationCell);
            if (destRow <= 0 || destCol <= 0) {
                throw new ArgumentException($"Invalid destination cell '{destinationCell}'.", nameof(destinationCell));
            }

            WriteLock(() => {
                var headers = BuildPivotHeaders(r1, c1, c2);
                if (headers.Count == 0) {
                    throw new InvalidOperationException("Pivot source range must include at least one header column.");
                }

                var calculatedFieldList = NormalizeCalculatedFields(calculatedFields, headers);
                var sourceHeaderIndex = BuildFieldIndex(headers);
                var groupingMap = BuildPivotGroupingMap(groupings, sourceHeaderIndex, headers.Count);
                var generatedGroupingFields = BuildGeneratedPivotGroupingFields(headers, groupingMap, calculatedFieldList);

                var allFields = new List<string>(headers);
                allFields.AddRange(generatedGroupingFields.Select(field => field.FieldName));
                allFields.AddRange(calculatedFieldList.Select(field => field.Name));

                var headerIndex = BuildFieldIndex(allFields);

                var dataFieldList = (dataFields ?? Array.Empty<ExcelPivotDataField>()).Where(df => df != null).ToList();
                if (dataFieldList.Count == 0) {
                    dataFieldList.Add(new ExcelPivotDataField(headers[headers.Count - 1], DataConsolidateFunctionValues.Sum));
                }
                var pivotFilterList = (pivotFilters ?? Array.Empty<ExcelPivotFilter>()).Where(filter => filter != null).ToList();
                var generatedFieldsBySource = BuildGeneratedPivotGroupingFieldMap(generatedGroupingFields);

                var rowFieldIndices = ResolveFieldIndices(rowFields, headerIndex, nameof(rowFields));
                var columnFieldIndices = ResolveFieldIndices(columnFields, headerIndex, nameof(columnFields));
                var pageFieldIndices = ResolveFieldIndices(pageFields, headerIndex, nameof(pageFields));
                if (pageFieldIndices.Count > 0) {
                    rowFieldIndices.RemoveAll(idx => pageFieldIndices.Contains(idx));
                    columnFieldIndices.RemoveAll(idx => pageFieldIndices.Contains(idx));
                }

                // If no row/column fields provided, default to the first non-data field when possible.
                if (rowFields == null && columnFields == null && rowFieldIndices.Count == 0 && columnFieldIndices.Count == 0) {
                    int dataIdx = ResolveFieldIndex(dataFieldList[0].FieldName, headerIndex, nameof(dataFields));
                    int fallback = headers.Count > 1 && dataIdx == 0 ? 1 : 0;
                    if (fallback >= 0 && fallback < headers.Count && fallback != dataIdx) {
                        rowFieldIndices.Add(fallback);
                    }
                }

                (rowFieldIndices, columnFieldIndices, pageFieldIndices) = ExpandGeneratedGroupingFieldIndices(
                    rowFieldIndices,
                    columnFieldIndices,
                    pageFieldIndices,
                    generatedFieldsBySource);

                var dataFieldIndices = new HashSet<int>();
                foreach (var df in dataFieldList) {
                    int idx = ResolveFieldIndex(df.FieldName, headerIndex, nameof(dataFields));
                    dataFieldIndices.Add(idx);
                }

                var workbookPart = WorkbookPartRoot;
                var workbook = workbookPart.Workbook ??= new Workbook();
                var fieldValueMap = BuildPivotFieldValueMap(headers.Count, r1 + 1, r2, c1, groupingMap);
                var generatedFieldValueMap = BuildGeneratedPivotFieldValueMap(generatedGroupingFields, r1 + 1, r2, c1);
                var allFieldValueMap = fieldValueMap
                    .Select(values => values.TextValues)
                    .Cast<IReadOnlyList<string>>()
                    .ToList();
                allFieldValueMap.AddRange(generatedFieldValueMap.Select(values => values.TextValues).Cast<IReadOnlyList<string>>());
                for (int i = 0; i < calculatedFieldList.Count; i++) {
                    allFieldValueMap.Add(Array.Empty<string>());
                }
                var fieldOptionMap = BuildPivotFieldOptionMap(fieldOptions, headerIndex);
                ExpandGeneratedGroupingFieldOptions(fieldOptionMap, generatedFieldsBySource, allFields, allFieldValueMap);
                uint cacheId = NextPivotCacheId(workbookPart);

                var cacheDefPart = workbookPart.AddNewPart<PivotTableCacheDefinitionPart>();
                var cacheDef = new PivotCacheDefinition {
                    CacheSource = new CacheSource {
                        Type = SourceValues.Worksheet,
                        WorksheetSource = new WorksheetSource {
                            Sheet = Name,
                            Reference = sourceRange
                        }
                    },
                    CacheFields = new CacheFields { Count = (uint)allFields.Count },
                    RecordCount = 0,
                    RefreshOnLoad = true,
                    SaveData = false
                };

                for (int i = 0; i < headers.Count; i++) {
                    string header = headers[i];
                    var cacheField = new CacheField { Name = header };
                    groupingMap.TryGetValue(i, out var grouping);
                    cacheField.SharedItems = BuildSharedItems(fieldValueMap[i], grouping);
                    if (grouping != null) {
                        cacheField.FieldGroup = CreatePivotFieldGroup(grouping, fieldValueMap[i]);
                    }
                    cacheDef.CacheFields.Append(cacheField);
                }

                for (int i = 0; i < generatedGroupingFields.Count; i++) {
                    var generatedField = generatedGroupingFields[i];
                    var cacheField = new CacheField {
                        Name = generatedField.FieldName,
                        DatabaseField = false
                    };
                    cacheField.SharedItems = BuildSharedItems(generatedFieldValueMap[i], generatedField.Grouping);
                    cacheField.FieldGroup = CreatePivotFieldGroup(
                        generatedField.Grouping,
                        generatedFieldValueMap[i],
                        (uint)generatedField.SourceIndex,
                        generatedField.ParentFieldIndex.HasValue ? (uint)generatedField.ParentFieldIndex.Value : null);
                    cacheDef.CacheFields.Append(cacheField);
                }

                foreach (var calculatedField in calculatedFieldList) {
                    var cacheField = new CacheField {
                        Name = calculatedField.Name,
                        Formula = calculatedField.Formula,
                        DatabaseField = false
                    };
                    if (!string.IsNullOrWhiteSpace(calculatedField.Caption)) cacheField.Caption = calculatedField.Caption;
                    uint? numberFormatId = ResolveNumberFormatId(workbookPart, calculatedField.NumberFormatId, calculatedField.NumberFormat);
                    if (numberFormatId.HasValue) cacheField.NumberFormatId = numberFormatId.Value;
                    cacheDef.CacheFields.Append(cacheField);
                }

                cacheDefPart.PivotCacheDefinition = cacheDef;
                cacheDefPart.PivotCacheDefinition.Save();

                var cacheRecordsPart = cacheDefPart.AddNewPart<PivotTableCacheRecordsPart>();
                cacheRecordsPart.PivotCacheRecords = new PivotCacheRecords { Count = 0U };
                cacheRecordsPart.PivotCacheRecords.Save();

                var pivotCaches = workbook.PivotCaches ?? workbook.AppendChild(new PivotCaches());
                pivotCaches.Append(new PivotCache {
                    CacheId = cacheId,
                    Id = workbookPart.GetIdOfPart(cacheDefPart)
                });
                // Count attribute is optional; OpenXml SDK does not expose a setter for PivotCaches.Count in all targets.

                var existingNames = _worksheetPart.PivotTableParts
                    .Select(p => p.PivotTableDefinition?.Name?.Value)
                    .Where(n => !string.IsNullOrWhiteSpace(n))
                    .Select(n => n!)
                    .ToList();
                string pivotName = EnsureUniquePivotTableName(name, existingNames);

                var pivotPart = _worksheetPart.AddNewPart<PivotTablePart>();
                pivotPart.AddPart(cacheDefPart);

                var pivotFields = new PivotFields { Count = (uint)allFields.Count };
                for (int i = 0; i < allFields.Count; i++) {
                    fieldOptionMap.TryGetValue(i, out var options);
                    var pivotField = new PivotField { ShowAll = options?.ShowAll ?? true };
                    if (pageFieldIndices.Contains(i)) pivotField.Axis = PivotTableAxisValues.AxisPage;
                    if (rowFieldIndices.Contains(i)) pivotField.Axis = PivotTableAxisValues.AxisRow;
                    if (columnFieldIndices.Contains(i)) pivotField.Axis = PivotTableAxisValues.AxisColumn;
                    if (dataFieldIndices.Contains(i)) pivotField.DataField = true;
                    ApplyPivotFieldOptions(pivotField, options, workbookPart, allFieldValueMap[i]);
                    pivotFields.Append(pivotField);
                }

                var rowFieldsElement = rowFieldIndices.Count > 0 ? new RowFields { Count = (uint)rowFieldIndices.Count } : null;
                if (rowFieldsElement != null) {
                    foreach (int idx in rowFieldIndices) rowFieldsElement.Append(new Field { Index = idx });
                }

                var columnFieldsElement = columnFieldIndices.Count > 0 ? new ColumnFields { Count = (uint)columnFieldIndices.Count } : null;
                if (columnFieldsElement != null) {
                    foreach (int idx in columnFieldIndices) columnFieldsElement.Append(new Field { Index = idx });
                }

                var pageFieldsElement = pageFieldIndices.Count > 0 ? new PageFields { Count = (uint)pageFieldIndices.Count } : null;
                if (pageFieldsElement != null) {
                    foreach (int idx in pageFieldIndices) {
                        fieldOptionMap.TryGetValue(idx, out var options);
                        pageFieldsElement.Append(CreatePageField(idx, options, allFieldValueMap[idx]));
                    }
                }

                var dataFieldsElement = new DataFields { Count = (uint)dataFieldList.Count };
                foreach (var df in dataFieldList) {
                    int idx = ResolveFieldIndex(df.FieldName, headerIndex, nameof(dataFields));
                    string display = df.DisplayName ?? $"{df.Function} of {allFields[idx]}";
                    var dataField = new DataField {
                        Name = display,
                        Field = (uint)idx,
                        Subtotal = df.Function
                    };
                    uint? numberFormatId = ResolveNumberFormatId(workbookPart, df.NumberFormatId, df.NumberFormat);
                    if (numberFormatId.HasValue) dataField.NumberFormatId = numberFormatId.Value;
                    if (df.ShowDataAs.HasValue) dataField.ShowDataAs = df.ShowDataAs.Value;
                    if (df.BaseField.HasValue) dataField.BaseField = df.BaseField.Value;
                    if (df.BaseItem.HasValue) dataField.BaseItem = df.BaseItem.Value;
                    dataFieldsElement.Append(dataField);
                }

                PivotFilters? pivotFiltersElement = CreatePivotFilters(pivotFilterList, headerIndex, dataFieldList);

                string pivotRef = BuildPivotLocationReference(destRow, destCol, rowFieldIndices.Count + columnFieldIndices.Count + dataFieldList.Count);

                var pivotDefinition = new PivotTableDefinition {
                    Name = pivotName,
                    CacheId = cacheId,
                    ApplyNumberFormats = true,
                    ApplyBorderFormats = true,
                    ApplyAlignmentFormats = true,
                    ApplyWidthHeightFormats = true,
                    ApplyPatternFormats = true,
                    UseAutoFormatting = true,
                    PreserveFormatting = true,
                    RowGrandTotals = showRowGrandTotals,
                    ColumnGrandTotals = showColumnGrandTotals,
                    MultipleFieldFilters = true,
                    DataCaption = "Values",
                    PivotTableStyleName = string.IsNullOrWhiteSpace(pivotStyleName) ? null : pivotStyleName,
                    Location = new Location {
                        Reference = pivotRef,
                        FirstHeaderRow = 1U,
                        FirstDataRow = 2U,
                        FirstDataColumn = 1U
                    },
                    PivotFields = pivotFields,
                    PivotFilters = pivotFiltersElement,
                    DataFields = dataFieldsElement
                };

                if (rowFieldsElement != null) pivotDefinition.RowFields = rowFieldsElement;
                if (columnFieldsElement != null) pivotDefinition.ColumnFields = columnFieldsElement;
                if (pageFieldsElement != null) pivotDefinition.PageFields = pageFieldsElement;

                switch (layout) {
                    case ExcelPivotLayout.Compact:
                        pivotDefinition.CompactData = true;
                        pivotDefinition.OutlineData = false;
                        break;
                    case ExcelPivotLayout.Outline:
                        pivotDefinition.CompactData = false;
                        pivotDefinition.OutlineData = true;
                        break;
                    case ExcelPivotLayout.Tabular:
                        pivotDefinition.CompactData = false;
                        pivotDefinition.OutlineData = false;
                        break;
                }

                if (dataOnRows.HasValue) pivotDefinition.DataOnRows = dataOnRows.Value;
                if (showHeaders.HasValue) pivotDefinition.ShowHeaders = showHeaders.Value;
                if (showEmptyRows.HasValue) pivotDefinition.ShowEmptyRow = showEmptyRows.Value;
                if (showEmptyColumns.HasValue) pivotDefinition.ShowEmptyColumn = showEmptyColumns.Value;
                if (showDrill.HasValue) pivotDefinition.ShowDrill = showDrill.Value;
                if (!string.IsNullOrWhiteSpace(rowHeaderCaption)) pivotDefinition.RowHeaderCaption = rowHeaderCaption;
                if (!string.IsNullOrWhiteSpace(columnHeaderCaption)) pivotDefinition.ColumnHeaderCaption = columnHeaderCaption;
                if (!string.IsNullOrWhiteSpace(grandTotalCaption)) pivotDefinition.GrandTotalCaption = grandTotalCaption;
                if (!string.IsNullOrWhiteSpace(missingCaption)) pivotDefinition.MissingCaption = missingCaption;
                if (!string.IsNullOrWhiteSpace(errorCaption)) pivotDefinition.ErrorCaption = errorCaption;
                if (showDataDropDown.HasValue) pivotDefinition.ShowDataDropDown = showDataDropDown.Value;
                if (showDropZones.HasValue) pivotDefinition.ShowDropZones = showDropZones.Value;
                if (showDataTips.HasValue) pivotDefinition.ShowDataTips = showDataTips.Value;
                if (showMemberPropertyTips.HasValue) pivotDefinition.ShowMemberPropertyTips = showMemberPropertyTips.Value;
                if (fieldListSortAscending.HasValue) pivotDefinition.FieldListSortAscending = fieldListSortAscending.Value;
                if (customListSort.HasValue) pivotDefinition.CustomListSort = customListSort.Value;

                pivotPart.PivotTableDefinition = pivotDefinition;
                pivotPart.PivotTableDefinition.Save();

                WorksheetRoot.Save();
                workbook.Save();
            });
        }

        private List<string> BuildPivotHeaders(int headerRow, int startColumn, int endColumn) {
            var headers = new List<string>();
            var used = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            for (int col = startColumn; col <= endColumn; col++) {
                string header = string.Empty;
                if (TryGetCellText(headerRow, col, out var text)) {
                    header = text?.Trim() ?? string.Empty;
                }
                if (string.IsNullOrWhiteSpace(header)) {
                    header = $"Column{col}";
                }
                header = EnsureUniqueName(header, used);
                used.Add(header);
                headers.Add(header);
            }
            return headers;
        }

        private static string EnsureUniqueName(string name, HashSet<string> used) {
            string baseName = string.IsNullOrWhiteSpace(name) ? "Column" : name.Trim();
            if (!used.Contains(baseName)) return baseName;
            int i = 2;
            string candidate;
            do {
                candidate = $"{baseName}_{i}";
                i++;
            } while (used.Contains(candidate));
            return candidate;
        }

        private static string EnsureUniquePivotTableName(string? name, IEnumerable<string> existingNames) {
            string baseName = string.IsNullOrWhiteSpace(name) ? "PivotTable" : name!.Trim();
            var existing = new HashSet<string>(existingNames, StringComparer.OrdinalIgnoreCase);
            if (!existing.Contains(baseName)) return baseName;
            int i = 2;
            string candidate;
            do {
                candidate = $"{baseName}{i}";
                i++;
            } while (existing.Contains(candidate));
            return candidate;
        }

        private static List<int> ResolveFieldIndices(IEnumerable<string>? fields, IDictionary<string, int> headerIndex, string paramName) {
            var indices = new List<int>();
            if (fields == null) return indices;
            foreach (var field in fields) {
                if (string.IsNullOrWhiteSpace(field)) continue;
                int idx = ResolveFieldIndex(field, headerIndex, paramName);
                if (!indices.Contains(idx)) indices.Add(idx);
            }
            return indices;
        }

        private static Dictionary<string, int> BuildFieldIndex(IReadOnlyList<string> fields) {
            var index = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            for (int i = 0; i < fields.Count; i++) {
                index[fields[i]] = i;
            }

            return index;
        }

        private static List<GeneratedPivotGroupingField> BuildGeneratedPivotGroupingFields(
            IReadOnlyList<string> headers,
            IReadOnlyDictionary<int, ExcelPivotGrouping> groupingMap,
            IReadOnlyList<ExcelPivotCalculatedField> calculatedFields) {
            var fields = new List<GeneratedPivotGroupingField>();
            var used = new HashSet<string>(headers, StringComparer.OrdinalIgnoreCase);
            foreach (var calculatedField in calculatedFields) {
                used.Add(calculatedField.Name);
            }

            foreach (var pair in groupingMap.OrderBy(pair => pair.Key)) {
                var grouping = pair.Value;
                if (!grouping.HasGeneratedDateLevels) continue;

                int? parentFieldIndex = null;
                foreach (var level in grouping.GeneratedDateLevels) {
                    string fieldName = EnsureUniqueName($"{headers[pair.Key]} {GetDateGroupFieldSuffix(level)}", used);
                    used.Add(fieldName);
                    var generatedGrouping = ExcelPivotGrouping.Date(fieldName, level, grouping.StartDate, grouping.EndDate);
                    int fieldIndex = headers.Count + fields.Count;
                    fields.Add(new GeneratedPivotGroupingField(pair.Key, fieldIndex, parentFieldIndex, fieldName, level, generatedGrouping));
                    parentFieldIndex = fieldIndex;
                }
            }

            return fields;
        }

        private static Dictionary<int, IReadOnlyList<int>> BuildGeneratedPivotGroupingFieldMap(IReadOnlyList<GeneratedPivotGroupingField> generatedFields) {
            var map = new Dictionary<int, IReadOnlyList<int>>();
            foreach (var group in generatedFields.GroupBy(field => field.SourceIndex)) {
                map[group.Key] = group.Select(field => field.FieldIndex).ToArray();
            }

            return map;
        }

        private static List<int> ExpandGeneratedGroupingFieldIndices(
            IReadOnlyList<int> fieldIndices,
            IReadOnlyDictionary<int, IReadOnlyList<int>> generatedFieldsBySource) {
            var expanded = new List<int>();
            foreach (int index in fieldIndices) {
                if (generatedFieldsBySource.TryGetValue(index, out var generatedFields)) {
                    foreach (int generatedIndex in generatedFields) {
                        if (!expanded.Contains(generatedIndex)) {
                            expanded.Add(generatedIndex);
                        }
                    }
                } else if (!expanded.Contains(index)) {
                    expanded.Add(index);
                }
            }

            return expanded;
        }

        private static (List<int> RowFields, List<int> ColumnFields, List<int> PageFields) ExpandGeneratedGroupingFieldIndices(
            IReadOnlyList<int> rowFieldIndices,
            IReadOnlyList<int> columnFieldIndices,
            IReadOnlyList<int> pageFieldIndices,
            IReadOnlyDictionary<int, IReadOnlyList<int>> generatedFieldsBySource) {
            return (
                ExpandGeneratedGroupingFieldIndices(rowFieldIndices, generatedFieldsBySource),
                ExpandGeneratedGroupingFieldIndices(columnFieldIndices, generatedFieldsBySource),
                ExpandGeneratedGroupingFieldIndices(pageFieldIndices, generatedFieldsBySource));
        }

        private static Dictionary<int, ExcelPivotFieldOptions> BuildPivotFieldOptionMap(IEnumerable<ExcelPivotFieldOptions>? fieldOptions,
            IDictionary<string, int> headerIndex) {
            var map = new Dictionary<int, ExcelPivotFieldOptions>();
            if (fieldOptions == null) return map;

            foreach (var options in fieldOptions) {
                if (options == null || string.IsNullOrWhiteSpace(options.FieldName)) continue;
                int idx = ResolveFieldIndex(options.FieldName, headerIndex, nameof(fieldOptions));
                map[idx] = options;
            }

            return map;
        }

        private static void ExpandGeneratedGroupingFieldOptions(
            IDictionary<int, ExcelPivotFieldOptions> fieldOptionMap,
            IReadOnlyDictionary<int, IReadOnlyList<int>> generatedFieldsBySource,
            IReadOnlyList<string> allFields,
            IReadOnlyList<IReadOnlyList<string>> allFieldValueMap) {
            foreach (var pair in generatedFieldsBySource) {
                if (!fieldOptionMap.TryGetValue(pair.Key, out var sourceOptions)) continue;

                fieldOptionMap[pair.Key] = ClonePivotFieldOptions(sourceOptions, sourceOptions.FieldName);
                foreach (int generatedIndex in pair.Value) {
                    if (generatedIndex < 0 || generatedIndex >= allFields.Count || generatedIndex >= allFieldValueMap.Count) continue;
                    fieldOptionMap[generatedIndex] = ClonePivotFieldOptionsForGeneratedField(
                        sourceOptions,
                        allFields[generatedIndex],
                        allFieldValueMap[generatedIndex]);
                }
            }
        }

        private static ExcelPivotFieldOptions ClonePivotFieldOptionsForGeneratedField(
            ExcelPivotFieldOptions sourceOptions,
            string fieldName,
            IReadOnlyList<string> generatedValues) {
            var valueSet = new HashSet<string>(generatedValues, StringComparer.OrdinalIgnoreCase);
            string[] hiddenItems = sourceOptions.HiddenItems.Where(valueSet.Contains).ToArray();
            string[] visibleItems = sourceOptions.VisibleItems.Where(valueSet.Contains).ToArray();
            string? selectedItem = sourceOptions.SelectedItem != null && valueSet.Contains(sourceOptions.SelectedItem)
                ? sourceOptions.SelectedItem
                : null;

            return ClonePivotFieldOptions(sourceOptions, fieldName, hiddenItems, visibleItems, selectedItem);
        }

        private static ExcelPivotFieldOptions ClonePivotFieldOptions(
            ExcelPivotFieldOptions sourceOptions,
            string fieldName,
            IEnumerable<string>? hiddenItems = null,
            IEnumerable<string>? visibleItems = null,
            string? selectedItem = null) {
            return new ExcelPivotFieldOptions(
                fieldName,
                sourceOptions.SortType,
                sourceOptions.NumberFormatId,
                sourceOptions.NumberFormat,
                sourceOptions.ShowAll,
                sourceOptions.DefaultSubtotal,
                sourceOptions.SubtotalTop,
                sourceOptions.InsertBlankRow,
                sourceOptions.InsertPageBreak,
                sourceOptions.Compact,
                sourceOptions.Outline,
                sourceOptions.ShowDropDowns,
                sourceOptions.MultipleItemSelectionAllowed,
                sourceOptions.IncludeNewItemsInFilter,
                sourceOptions.SubtotalCaption,
                hiddenItems,
                visibleItems,
                selectedItem);
        }

        private static Dictionary<int, ExcelPivotGrouping> BuildPivotGroupingMap(IEnumerable<ExcelPivotGrouping>? groupings,
            IDictionary<string, int> headerIndex,
            int sourceFieldCount) {
            var map = new Dictionary<int, ExcelPivotGrouping>();
            if (groupings == null) return map;

            foreach (var grouping in groupings) {
                if (grouping == null) continue;
                int idx = ResolveFieldIndex(grouping.FieldName, headerIndex, nameof(groupings));
                if (idx >= sourceFieldCount) {
                    throw new ArgumentException($"Pivot grouping field '{grouping.FieldName}' must be a source field, not a calculated field.", nameof(groupings));
                }
                map[idx] = grouping;
            }

            return map;
        }

        private static List<ExcelPivotCalculatedField> NormalizeCalculatedFields(IEnumerable<ExcelPivotCalculatedField>? calculatedFields,
            IReadOnlyList<string> sourceHeaders) {
            var list = new List<ExcelPivotCalculatedField>();
            if (calculatedFields == null) return list;

            var names = new HashSet<string>(sourceHeaders, StringComparer.OrdinalIgnoreCase);
            foreach (var field in calculatedFields) {
                if (field == null) continue;
                if (!names.Add(field.Name)) {
                    throw new ArgumentException($"Pivot calculated field '{field.Name}' duplicates an existing source or calculated field name.", nameof(calculatedFields));
                }

                list.Add(field);
            }

            return list;
        }

        private List<PivotFieldValues> BuildPivotFieldValueMap(int fieldCount, int firstDataRow, int lastDataRow, int firstColumn,
            IReadOnlyDictionary<int, ExcelPivotGrouping> groupingMap) {
            var maps = new List<PivotFieldValues>(fieldCount);
            for (int field = 0; field < fieldCount; field++) {
                var values = new List<PivotFieldValue>();
                var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                int column = firstColumn + field;
                groupingMap.TryGetValue(field, out var grouping);
                for (int row = firstDataRow; row <= lastDataRow; row++) {
                    var value = GetPivotFieldValue(row, column, grouping);
                    string text = value.Text;
                    if (seen.Add(text)) {
                        values.Add(value);
                    }
                }

                maps.Add(new PivotFieldValues(values));
            }

            return maps;
        }

        private List<PivotFieldValues> BuildGeneratedPivotFieldValueMap(
            IReadOnlyList<GeneratedPivotGroupingField> generatedFields,
            int firstDataRow,
            int lastDataRow,
            int firstColumn) {
            var maps = new List<PivotFieldValues>(generatedFields.Count);
            foreach (var generatedField in generatedFields) {
                var values = new List<PivotFieldValue>();
                var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                int column = firstColumn + generatedField.SourceIndex;
                for (int row = firstDataRow; row <= lastDataRow; row++) {
                    var value = GetGeneratedPivotDateFieldValue(row, column, generatedField.GroupBy);
                    if (seen.Add(value.Text)) {
                        values.Add(value);
                    }
                }

                maps.Add(new PivotFieldValues(values));
            }

            return maps;
        }

        private PivotFieldValue GetPivotFieldValue(int row, int column, ExcelPivotGrouping? grouping) {
            string text = TryGetCellText(row, column, out string cellText) ? cellText.Trim() : string.Empty;
            if (string.IsNullOrEmpty(text)) {
                return PivotFieldValue.Blank();
            }

            if (grouping?.IsDateGrouping == true) {
                if (TryGetPivotDateValue(row, column, text, out var date)) {
                    return PivotFieldValue.FromDate(date);
                }
            }

            if (grouping?.GroupBy == GroupByValues.Range) {
                var snapshot = GetCellValueSnapshot(row, column);
                if (snapshot.Value is double number) {
                    return PivotFieldValue.FromNumber(number);
                }

                if (double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out number)) {
                    return PivotFieldValue.FromNumber(number);
                }
            }

            return PivotFieldValue.FromText(text);
        }

        private PivotFieldValue GetGeneratedPivotDateFieldValue(int row, int column, GroupByValues groupBy) {
            string text = TryGetCellText(row, column, out string cellText) ? cellText.Trim() : string.Empty;
            if (string.IsNullOrEmpty(text)) {
                return PivotFieldValue.Blank();
            }

            return TryGetPivotDateValue(row, column, text, out var date)
                ? PivotFieldValue.FromText(FormatGeneratedDateGroupValue(date, groupBy))
                : PivotFieldValue.FromText(text);
        }

        private bool TryGetPivotDateValue(int row, int column, string text, out DateTime date) {
            var snapshot = GetCellValueSnapshot(row, column);
            if (snapshot.Value is double serial) {
                try {
                    date = DateTime.FromOADate(serial);
                    return true;
                } catch {
                    // Fall through to string parsing when a numeric value is not a valid Excel date.
                }
            }

            if (DateTime.TryParse(text, CultureInfo.CurrentCulture, DateTimeStyles.None, out date)
                || DateTime.TryParse(text, CultureInfo.InvariantCulture, DateTimeStyles.None, out date)) {
                return true;
            }

            date = default;
            return false;
        }

        private static string FormatGeneratedDateGroupValue(DateTime date, GroupByValues groupBy) {
            if (groupBy == GroupByValues.Years) return date.Year.ToString(CultureInfo.InvariantCulture);
            if (groupBy == GroupByValues.Quarters) return $"Q{((date.Month - 1) / 3) + 1}";
            if (groupBy == GroupByValues.Months) return date.ToString("MMMM", CultureInfo.InvariantCulture);
            if (groupBy == GroupByValues.Days) return date.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
            if (groupBy == GroupByValues.Hours) return date.Hour.ToString("00", CultureInfo.InvariantCulture);
            if (groupBy == GroupByValues.Minutes) return date.ToString("HH:mm", CultureInfo.InvariantCulture);
            if (groupBy == GroupByValues.Seconds) return date.ToString("HH:mm:ss", CultureInfo.InvariantCulture);
            return date.ToString("O", CultureInfo.InvariantCulture);
        }

        private static string GetDateGroupFieldSuffix(GroupByValues groupBy) {
            if (groupBy == GroupByValues.Years) return "Years";
            if (groupBy == GroupByValues.Quarters) return "Quarters";
            if (groupBy == GroupByValues.Months) return "Months";
            if (groupBy == GroupByValues.Days) return "Days";
            if (groupBy == GroupByValues.Hours) return "Hours";
            if (groupBy == GroupByValues.Minutes) return "Minutes";
            if (groupBy == GroupByValues.Seconds) return "Seconds";
            return groupBy.ToString();
        }

        private static SharedItems BuildSharedItems(PivotFieldValues values, ExcelPivotGrouping? grouping) {
            bool hasBlank = values.Items.Any(value => value.Kind == PivotFieldValueKind.Blank);
            bool hasDate = values.Items.Any(value => value.Kind == PivotFieldValueKind.Date);
            bool hasNumber = values.Items.Any(value => value.Kind == PivotFieldValueKind.Number);
            bool hasString = values.Items.Any(value => value.Kind == PivotFieldValueKind.Text);
            var sharedItems = new SharedItems {
                Count = (uint)values.Items.Count,
                ContainsString = hasString,
                ContainsBlank = hasBlank,
                ContainsDate = hasDate,
                ContainsNumber = hasNumber
            };

            if (hasNumber) {
                var numericValues = values.Items.Where(value => value.Number.HasValue).Select(value => value.Number!.Value).ToList();
                if (numericValues.Count > 0) {
                    sharedItems.MinValue = numericValues.Min();
                    sharedItems.MaxValue = numericValues.Max();
                    sharedItems.ContainsInteger = numericValues.All(value => Math.Abs(value - Math.Round(value)) < 0.0000001d);
                }
            }

            if (hasDate) {
                var dateValues = values.Items.Where(value => value.Date.HasValue).Select(value => value.Date!.Value).ToList();
                if (dateValues.Count > 0) {
                    sharedItems.MinDate = dateValues.Min();
                    sharedItems.MaxDate = dateValues.Max();
                }
            }

            foreach (var value in values.Items) {
                sharedItems.Append(value.Kind switch {
                    PivotFieldValueKind.Blank => new MissingItem(),
                    PivotFieldValueKind.Number => new NumberItem { Val = value.Number!.Value },
                    PivotFieldValueKind.Date => new DateTimeItem { Val = value.Date!.Value },
                    _ => new StringItem { Val = value.Text }
                });
            }

            return sharedItems;
        }

        private static FieldGroup CreatePivotFieldGroup(ExcelPivotGrouping grouping, PivotFieldValues? groupItems = null, uint? baseFieldIndex = null, uint? parentFieldIndex = null) {
            var range = new RangeProperties {
                AutoStart = grouping.AutoStart,
                AutoEnd = grouping.AutoEnd,
                GroupBy = grouping.GroupBy
            };

            if (grouping.StartDate.HasValue) range.StartDate = grouping.StartDate.Value;
            if (grouping.EndDate.HasValue) range.EndDate = grouping.EndDate.Value;
            if (grouping.StartNumber.HasValue) range.StartNumber = grouping.StartNumber.Value;
            if (grouping.EndNumber.HasValue) range.EndNum = grouping.EndNumber.Value;
            if (grouping.Interval.HasValue) range.GroupInterval = grouping.Interval.Value;

            var fieldGroup = new FieldGroup(range);
            if (baseFieldIndex.HasValue) fieldGroup.Base = baseFieldIndex.Value;
            if (parentFieldIndex.HasValue) fieldGroup.ParentId = parentFieldIndex.Value;
            if (groupItems != null) {
                fieldGroup.Append(BuildGroupItems(groupItems));
            }

            return fieldGroup;
        }

        private static GroupItems BuildGroupItems(PivotFieldValues values) {
            var groupItems = new GroupItems { Count = (uint)values.Items.Count };
            foreach (var value in values.Items) {
                groupItems.Append(value.Kind switch {
                    PivotFieldValueKind.Blank => new MissingItem(),
                    PivotFieldValueKind.Number => new NumberItem { Val = value.Number!.Value },
                    PivotFieldValueKind.Date => new DateTimeItem { Val = value.Date!.Value },
                    _ => new StringItem { Val = value.Text }
                });
            }

            return groupItems;
        }

        private enum PivotFieldValueKind {
            Blank,
            Text,
            Number,
            Date
        }

        private sealed class PivotFieldValue {
            private PivotFieldValue(PivotFieldValueKind kind, string text, double? number = null, DateTime? date = null) {
                Kind = kind;
                Text = text;
                Number = number;
                Date = date;
            }

            public PivotFieldValueKind Kind { get; }

            public string Text { get; }

            public double? Number { get; }

            public DateTime? Date { get; }

            public static PivotFieldValue Blank() => new(PivotFieldValueKind.Blank, string.Empty);

            public static PivotFieldValue FromText(string text) => new(PivotFieldValueKind.Text, text);

            public static PivotFieldValue FromNumber(double number) => new(PivotFieldValueKind.Number, number.ToString("G17", CultureInfo.InvariantCulture), number);

            public static PivotFieldValue FromDate(DateTime date) => new(PivotFieldValueKind.Date, date.ToString("O", CultureInfo.InvariantCulture), date: date);
        }

        private sealed class PivotFieldValues {
            public PivotFieldValues(IReadOnlyList<PivotFieldValue> items) {
                Items = items;
                TextValues = items.Select(item => item.Text).ToArray();
            }

            public IReadOnlyList<PivotFieldValue> Items { get; }

            public IReadOnlyList<string> TextValues { get; }
        }

        private sealed class GeneratedPivotGroupingField {
            public GeneratedPivotGroupingField(int sourceIndex, int fieldIndex, int? parentFieldIndex, string fieldName, GroupByValues groupBy, ExcelPivotGrouping grouping) {
                SourceIndex = sourceIndex;
                FieldIndex = fieldIndex;
                ParentFieldIndex = parentFieldIndex;
                FieldName = fieldName;
                GroupBy = groupBy;
                Grouping = grouping;
            }

            public int SourceIndex { get; }

            public int FieldIndex { get; }

            public int? ParentFieldIndex { get; }

            public string FieldName { get; }

            public GroupByValues GroupBy { get; }

            public ExcelPivotGrouping Grouping { get; }
        }

        private static PageField CreatePageField(int fieldIndex, ExcelPivotFieldOptions? options, IReadOnlyList<string> values) {
            var pageField = new PageField { Field = fieldIndex };
            if (options == null || string.IsNullOrWhiteSpace(options.SelectedItem)) {
                return pageField;
            }

            int selectedIndex = FindPivotItemIndex(options.SelectedItem!, values, options.FieldName, nameof(options.SelectedItem));
            pageField.Item = (uint)selectedIndex;
            return pageField;
        }

        private static void ApplyPivotFieldOptions(PivotField pivotField, ExcelPivotFieldOptions? options, WorkbookPart workbookPart,
            IReadOnlyList<string> values) {
            if (options == null) return;

            if (options.SortType.HasValue) pivotField.SortType = options.SortType.Value;
            if (options.DefaultSubtotal.HasValue) pivotField.DefaultSubtotal = options.DefaultSubtotal.Value;
            if (options.SubtotalTop.HasValue) pivotField.SubtotalTop = options.SubtotalTop.Value;
            if (options.InsertBlankRow.HasValue) pivotField.InsertBlankRow = options.InsertBlankRow.Value;
            if (options.InsertPageBreak.HasValue) pivotField.InsertPageBreak = options.InsertPageBreak.Value;
            if (options.Compact.HasValue) pivotField.Compact = options.Compact.Value;
            if (options.Outline.HasValue) pivotField.Outline = options.Outline.Value;
            if (options.ShowDropDowns.HasValue) pivotField.ShowDropDowns = options.ShowDropDowns.Value;
            if (options.MultipleItemSelectionAllowed.HasValue) pivotField.MultipleItemSelectionAllowed = options.MultipleItemSelectionAllowed.Value;
            if (options.IncludeNewItemsInFilter.HasValue) pivotField.IncludeNewItemsInFilter = options.IncludeNewItemsInFilter.Value;
            if (!string.IsNullOrWhiteSpace(options.SubtotalCaption)) pivotField.SubtotalCaption = options.SubtotalCaption;

            uint? numberFormatId = ResolveNumberFormatId(workbookPart, options.NumberFormatId, options.NumberFormat);
            if (numberFormatId.HasValue) pivotField.NumberFormatId = numberFormatId.Value;

            ApplyPivotFieldItemFilters(pivotField, options, values);
        }

        private static void ApplyPivotFieldItemFilters(PivotField pivotField, ExcelPivotFieldOptions options, IReadOnlyList<string> values) {
            if (options.HiddenItems.Count == 0 && options.VisibleItems.Count == 0) return;
            if (values.Count == 0) {
                throw new ArgumentException($"Field '{options.FieldName}' has no cache items to filter.", nameof(options));
            }

            var hidden = new HashSet<int>();
            if (options.HiddenItems.Count > 0) {
                foreach (string item in options.HiddenItems) {
                    hidden.Add(FindPivotItemIndex(item, values, options.FieldName, nameof(options.HiddenItems)));
                }
            } else {
                var visible = new HashSet<int>();
                foreach (string item in options.VisibleItems) {
                    visible.Add(FindPivotItemIndex(item, values, options.FieldName, nameof(options.VisibleItems)));
                }

                for (int i = 0; i < values.Count; i++) {
                    if (!visible.Contains(i)) hidden.Add(i);
                }
            }

            var items = new Items { Count = (uint)values.Count };
            for (int i = 0; i < values.Count; i++) {
                var item = new Item { Index = (uint)i };
                if (hidden.Contains(i)) item.Hidden = true;
                items.Append(item);
            }

            pivotField.Items = items;
            if (options.ShowAll == null) {
                pivotField.ShowAll = false;
            }
        }

        private static PivotFilters? CreatePivotFilters(IReadOnlyList<ExcelPivotFilter> filters,
            IDictionary<string, int> headerIndex,
            IReadOnlyList<ExcelPivotDataField> dataFields) {
            if (filters.Count == 0) return null;

            var pivotFilters = new PivotFilters { Count = (uint)filters.Count };
            for (int i = 0; i < filters.Count; i++) {
                var filter = filters[i];
                int fieldIndex = ResolveFieldIndex(filter.FieldName, headerIndex, nameof(filters));
                var pivotFilter = new PivotFilter {
                    Field = (uint)fieldIndex,
                    Type = filter.Type,
                    EvaluationOrder = i,
                    Id = (uint)(i + 1)
                };

                if (!string.IsNullOrWhiteSpace(filter.Value1)) pivotFilter.StringValue1 = filter.Value1;
                if (!string.IsNullOrWhiteSpace(filter.Value2)) pivotFilter.StringValue2 = filter.Value2;
                if (!string.IsNullOrWhiteSpace(filter.Name)) pivotFilter.Name = filter.Name;
                if (!string.IsNullOrWhiteSpace(filter.Description)) pivotFilter.Description = filter.Description;
                if (!string.IsNullOrWhiteSpace(filter.DataFieldName)) {
                    pivotFilter.MeasureField = (uint)ResolvePivotDataFieldIndex(filter.DataFieldName!, dataFields);
                }

                pivotFilter.AutoFilter = CreatePivotFilterAutoFilter(filter);
                pivotFilters.Append(pivotFilter);
            }

            return pivotFilters;
        }

        private static AutoFilter CreatePivotFilterAutoFilter(ExcelPivotFilter filter) {
            var autoFilter = new AutoFilter { Reference = "A1" };
            var filterColumn = new FilterColumn { ColumnId = 0U };

            if (TryCreateTop10Filter(filter, out var top10)) {
                filterColumn.Append(top10);
                autoFilter.Append(filterColumn);
                return autoFilter;
            }

            if (TryCreateDynamicFilter(filter, out var dynamicFilter)) {
                filterColumn.Append(dynamicFilter);
                autoFilter.Append(filterColumn);
                return autoFilter;
            }

            CustomFilters customFilters;

            if (TryResolveBetweenFilter(filter.Type, out var firstOperator, out var secondOperator, out bool matchAll)) {
                if (filter.Value1 == null || filter.Value2 == null) {
                    throw new ArgumentException($"Pivot filter '{filter.Type}' requires two values.", nameof(filter));
                }

                customFilters = new CustomFilters { And = matchAll };
                customFilters.Append(new CustomFilter {
                    Operator = firstOperator,
                    Val = filter.Value1
                });
                customFilters.Append(new CustomFilter {
                    Operator = secondOperator,
                    Val = filter.Value2
                });
            } else {
                if (filter.Value1 == null) {
                    throw new ArgumentException($"Pivot filter '{filter.Type}' requires a value.", nameof(filter));
                }

                customFilters = new CustomFilters();
                customFilters.Append(new CustomFilter {
                    Operator = ResolveSingleFilterOperator(filter.Type),
                    Val = NormalizePivotFilterAutoFilterValue(filter.Type, filter.Value1)
                });
            }

            filterColumn.Append(customFilters);
            autoFilter.Append(filterColumn);
            return autoFilter;
        }

        private static bool TryCreateTop10Filter(ExcelPivotFilter filter, out Top10 top10) {
            if (filter.Type != PivotFilterValues.Count && filter.Type != PivotFilterValues.Percent && filter.Type != PivotFilterValues.Sum) {
                top10 = new Top10();
                return false;
            }

            if (filter.Value1 == null) {
                throw new ArgumentException($"Pivot filter '{filter.Type}' requires a value.", nameof(filter));
            }

            if (!double.TryParse(filter.Value1, NumberStyles.Float, CultureInfo.InvariantCulture, out double value)) {
                throw new ArgumentException($"Pivot filter '{filter.Type}' value '{filter.Value1}' is not numeric.", nameof(filter));
            }

            top10 = new Top10 {
                Top = filter.IsTop ?? true,
                Percent = filter.IsPercent ?? filter.Type == PivotFilterValues.Percent,
                Val = value
            };

            if (!string.IsNullOrWhiteSpace(filter.FilterValue)) {
                if (!double.TryParse(filter.FilterValue, NumberStyles.Float, CultureInfo.InvariantCulture, out double filterValue)) {
                    throw new ArgumentException($"Pivot filter '{filter.Type}' filter value '{filter.FilterValue}' is not numeric.", nameof(filter));
                }

                top10.FilterValue = filterValue;
            }

            return true;
        }

        private static bool TryCreateDynamicFilter(ExcelPivotFilter filter, out DynamicFilter dynamicFilter) {
            DynamicFilterValues? dynamicType = ResolveDynamicFilterType(filter.Type);
            if (!dynamicType.HasValue) {
                dynamicFilter = new DynamicFilter();
                return false;
            }

            dynamicFilter = new DynamicFilter { Type = dynamicType.Value };
            return true;
        }

        private static DynamicFilterValues? ResolveDynamicFilterType(PivotFilterValues type) {
            if (type == PivotFilterValues.Today) return DynamicFilterValues.Today;
            if (type == PivotFilterValues.Yesterday) return DynamicFilterValues.Yesterday;
            if (type == PivotFilterValues.Tomorrow) return DynamicFilterValues.Tomorrow;
            if (type == PivotFilterValues.ThisWeek) return DynamicFilterValues.ThisWeek;
            if (type == PivotFilterValues.LastWeek) return DynamicFilterValues.LastWeek;
            if (type == PivotFilterValues.NextWeek) return DynamicFilterValues.NextWeek;
            if (type == PivotFilterValues.ThisMonth) return DynamicFilterValues.ThisMonth;
            if (type == PivotFilterValues.LastMonth) return DynamicFilterValues.LastMonth;
            if (type == PivotFilterValues.NextMonth) return DynamicFilterValues.NextMonth;
            if (type == PivotFilterValues.ThisQuarter) return DynamicFilterValues.ThisQuarter;
            if (type == PivotFilterValues.LastQuarter) return DynamicFilterValues.LastQuarter;
            if (type == PivotFilterValues.NextQuarter) return DynamicFilterValues.NextQuarter;
            if (type == PivotFilterValues.ThisYear) return DynamicFilterValues.ThisYear;
            if (type == PivotFilterValues.LastYear) return DynamicFilterValues.LastYear;
            if (type == PivotFilterValues.NextYear) return DynamicFilterValues.NextYear;
            if (type == PivotFilterValues.YearToDate) return DynamicFilterValues.YearToDate;
            if (type == PivotFilterValues.January) return DynamicFilterValues.January;
            if (type == PivotFilterValues.February) return DynamicFilterValues.February;
            if (type == PivotFilterValues.March) return DynamicFilterValues.March;
            if (type == PivotFilterValues.April) return DynamicFilterValues.April;
            if (type == PivotFilterValues.May) return DynamicFilterValues.May;
            if (type == PivotFilterValues.June) return DynamicFilterValues.June;
            if (type == PivotFilterValues.July) return DynamicFilterValues.July;
            if (type == PivotFilterValues.August) return DynamicFilterValues.August;
            if (type == PivotFilterValues.September) return DynamicFilterValues.September;
            if (type == PivotFilterValues.October) return DynamicFilterValues.October;
            if (type == PivotFilterValues.November) return DynamicFilterValues.November;
            if (type == PivotFilterValues.December) return DynamicFilterValues.December;
            if (type == PivotFilterValues.Quarter1) return DynamicFilterValues.Quarter1;
            if (type == PivotFilterValues.Quarter2) return DynamicFilterValues.Quarter2;
            if (type == PivotFilterValues.Quarter3) return DynamicFilterValues.Quarter3;
            if (type == PivotFilterValues.Quarter4) return DynamicFilterValues.Quarter4;

            return null;
        }

        private static bool TryResolveBetweenFilter(PivotFilterValues type,
            out FilterOperatorValues firstOperator,
            out FilterOperatorValues secondOperator,
            out bool matchAll) {
            if (type == PivotFilterValues.CaptionBetween || type == PivotFilterValues.ValueBetween || type == PivotFilterValues.DateBetween) {
                firstOperator = FilterOperatorValues.GreaterThanOrEqual;
                secondOperator = FilterOperatorValues.LessThanOrEqual;
                matchAll = true;
                return true;
            }

            if (type == PivotFilterValues.CaptionNotBetween || type == PivotFilterValues.ValueNotBetween || type == PivotFilterValues.DateNotBetween) {
                firstOperator = FilterOperatorValues.LessThan;
                secondOperator = FilterOperatorValues.GreaterThan;
                matchAll = false;
                return true;
            }

            firstOperator = FilterOperatorValues.Equal;
            secondOperator = FilterOperatorValues.Equal;
            matchAll = true;
            return false;
        }

        private static FilterOperatorValues ResolveSingleFilterOperator(PivotFilterValues type) {
            if (type == PivotFilterValues.CaptionNotEqual || type == PivotFilterValues.CaptionNotContains
                || type == PivotFilterValues.CaptionNotBeginsWith || type == PivotFilterValues.CaptionNotEndsWith
                || type == PivotFilterValues.ValueNotEqual || type == PivotFilterValues.DateNotEqual) {
                return FilterOperatorValues.NotEqual;
            }

            if (type == PivotFilterValues.CaptionGreaterThan || type == PivotFilterValues.ValueGreaterThan || type == PivotFilterValues.DateNewerThan) {
                return FilterOperatorValues.GreaterThan;
            }

            if (type == PivotFilterValues.CaptionGreaterThanOrEqual || type == PivotFilterValues.ValueGreaterThanOrEqual || type == PivotFilterValues.DateNewerThanOrEqual) {
                return FilterOperatorValues.GreaterThanOrEqual;
            }

            if (type == PivotFilterValues.CaptionLessThan || type == PivotFilterValues.ValueLessThan || type == PivotFilterValues.DateOlderThan) {
                return FilterOperatorValues.LessThan;
            }

            if (type == PivotFilterValues.CaptionLessThanOrEqual || type == PivotFilterValues.ValueLessThanOrEqual || type == PivotFilterValues.DateOlderThanOrEqual) {
                return FilterOperatorValues.LessThanOrEqual;
            }

            return FilterOperatorValues.Equal;
        }

        private static string NormalizePivotFilterAutoFilterValue(PivotFilterValues type, string value) {
            if (type == PivotFilterValues.CaptionContains || type == PivotFilterValues.CaptionNotContains) {
                return "*" + value + "*";
            }

            if (type == PivotFilterValues.CaptionBeginsWith || type == PivotFilterValues.CaptionNotBeginsWith) {
                return value + "*";
            }

            if (type == PivotFilterValues.CaptionEndsWith || type == PivotFilterValues.CaptionNotEndsWith) {
                return "*" + value;
            }

            return value;
        }

        private static int ResolvePivotDataFieldIndex(string dataFieldName, IReadOnlyList<ExcelPivotDataField> dataFields) {
            for (int i = 0; i < dataFields.Count; i++) {
                var dataField = dataFields[i];
                if (string.Equals(dataField.FieldName, dataFieldName, StringComparison.OrdinalIgnoreCase)
                    || string.Equals(dataField.DisplayName, dataFieldName, StringComparison.OrdinalIgnoreCase)) {
                    return i;
                }
            }

            throw new ArgumentException($"Data field '{dataFieldName}' was not found in pivot data fields.", nameof(dataFieldName));
        }

        private static int FindPivotItemIndex(string item, IReadOnlyList<string> values, string fieldName, string paramName) {
            for (int i = 0; i < values.Count; i++) {
                if (string.Equals(values[i], item, StringComparison.OrdinalIgnoreCase)) {
                    return i;
                }
            }

            throw new ArgumentException($"Item '{item}' was not found in pivot field '{fieldName}'.", paramName);
        }

        private static uint? ResolveNumberFormatId(WorkbookPart workbookPart, uint? numberFormatId, string? numberFormat) {
            if (numberFormatId.HasValue) return numberFormatId.Value;
            if (string.IsNullOrWhiteSpace(numberFormat)) return null;
            return GetOrCreateNumberFormatId(workbookPart, numberFormat!.Trim());
        }

        private static uint GetOrCreateNumberFormatId(WorkbookPart workbookPart, string numberFormat) {
            WorkbookStylesPart? stylesPart = workbookPart.WorkbookStylesPart;
            if (stylesPart == null) {
                stylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
            }

            Stylesheet stylesheet = stylesPart.Stylesheet ??= new Stylesheet();
            EnsureDefaultStylePrimitives(stylesheet);

            stylesheet.NumberingFormats ??= new NumberingFormats();
            NumberingFormat? existingFormat = stylesheet.NumberingFormats.Elements<NumberingFormat>()
                .FirstOrDefault(n => n.FormatCode != null && string.Equals(n.FormatCode.Value, numberFormat, StringComparison.Ordinal));

            if (existingFormat?.NumberFormatId?.Value is uint existingId) {
                return existingId;
            }

            uint formatId = stylesheet.NumberingFormats.Elements<NumberingFormat>().Any()
                ? Math.Max(164U, stylesheet.NumberingFormats.Elements<NumberingFormat>().Max(n => n.NumberFormatId?.Value ?? 0U) + 1U)
                : 164U;

            stylesheet.NumberingFormats.Append(new NumberingFormat {
                NumberFormatId = formatId,
                FormatCode = StringValue.FromString(numberFormat)
            });
            stylesheet.NumberingFormats.Count = (uint)stylesheet.NumberingFormats.Count();
            stylesPart.Stylesheet.Save();
            return formatId;
        }

        private static int ResolveFieldIndex(string field, IDictionary<string, int> headerIndex, string paramName) {
            var key = field.Trim();
            if (!headerIndex.TryGetValue(key, out int idx)) {
                throw new ArgumentException($"Field '{field}' was not found in pivot source headers.", paramName);
            }
            return idx;
        }

        private static string BuildPivotLocationReference(int startRow, int startColumn, int columnCount) {
            int width = Math.Max(1, columnCount);
            int endColumn = startColumn + width - 1;
            int endRow = startRow + 1; // header + at least one data row
            string start = A1.CellReference(startRow, startColumn);
            string end = A1.CellReference(endRow, endColumn);
            return $"{start}:{end}";
        }

        private static uint NextPivotCacheId(WorkbookPart workbookPart) {
            var workbook = workbookPart.Workbook ?? throw new InvalidOperationException("Workbook is missing.");
            var pivotCaches = workbook.PivotCaches;
            if (pivotCaches == null) return 1;
            uint max = 0;
            foreach (var cache in pivotCaches.Elements<PivotCache>()) {
                if (cache.CacheId != null && cache.CacheId.Value > max) max = cache.CacheId.Value;
            }
            return max + 1;
        }

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
            foreach (var builtInFormat in BuiltInNumberFormatCodes) {
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
