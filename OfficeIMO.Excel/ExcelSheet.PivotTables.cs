using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Text;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {

        private static readonly IReadOnlyDictionary<int, IReadOnlyList<int>> EmptyGeneratedPivotGroupingFieldMap =
            new Dictionary<int, IReadOnlyList<int>>(0);

        private StylesCache? _pivotStylesCache;

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
                        groupings: groupingInfos,
                        refreshOnOpen: cacheDef?.RefreshOnLoad?.Value,
                        saveSourceData: cacheDef?.SaveData?.Value,
                        preserveFormatting: def.PreserveFormatting?.Value,
                        enableDrill: def.EnableDrill?.Value));
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
                groupings: null,
                options: null);
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
        /// <param name="options">Optional pivot cache and workbook-interaction settings.</param>
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
            IEnumerable<ExcelPivotGrouping>? groupings = null,
            ExcelPivotTableOptions? options = null) {
            if (string.IsNullOrWhiteSpace(sourceRange)) throw new ArgumentNullException(nameof(sourceRange));
            if (string.IsNullOrWhiteSpace(destinationCell)) throw new ArgumentNullException(nameof(destinationCell));
            if (!A1.TryParseRange(sourceRange, out int r1, out int c1, out int r2, out int c2)) {
                throw new ArgumentException($"Invalid A1 range '{sourceRange}'.", nameof(sourceRange));
            }

            var (destRow, destCol) = A1.ParseCellRef(destinationCell);
            if (destRow <= 0 || destCol <= 0) {
                throw new ArgumentException($"Invalid destination cell '{destinationCell}'.", nameof(destinationCell));
            }

            _excelDocument.TryGetDeferredDirectTabularPivotSource(this, r1, c1, r2, c2, out var deferredPivotSource);

            WriteLockWorksheetPreparationOnly(() => {
                Stopwatch? pivotWatch = EffectiveExecution.OnTiming == null ? null : Stopwatch.StartNew();
                void ReportPivotTiming(string operation) {
                    if (pivotWatch == null) {
                        return;
                    }

                    EffectiveExecution.ReportTiming(operation, pivotWatch.Elapsed);
                    pivotWatch.Restart();
                }

                var headers = deferredPivotSource != null
                    ? BuildPivotHeaders(deferredPivotSource, c1, c2)
                    : BuildPivotHeaders(r1, c1, c2);
                if (headers.Count == 0) {
                    throw new InvalidOperationException("Pivot source range must include at least one header column.");
                }
                ReportPivotTiming("AddPivotTable.BuildHeaders");

                var calculatedFieldList = NormalizeCalculatedFields(calculatedFields, headers);
                var sourceHeaderIndex = BuildFieldIndex(headers);
                var groupingMap = BuildPivotGroupingMap(groupings, sourceHeaderIndex, headers.Count);
                var generatedGroupingFields = BuildGeneratedPivotGroupingFields(headers, groupingMap, calculatedFieldList);

                IReadOnlyList<string> allFields = headers;
                var headerIndex = sourceHeaderIndex;
                if (generatedGroupingFields.Count != 0 || calculatedFieldList.Count != 0) {
                    var allFieldList = new List<string>(headers.Count + generatedGroupingFields.Count + calculatedFieldList.Count);
                    for (int i = 0; i < headers.Count; i++) {
                        allFieldList.Add(headers[i]);
                    }

                    for (int i = 0; i < generatedGroupingFields.Count; i++) {
                        allFieldList.Add(generatedGroupingFields[i].FieldName);
                    }

                    for (int i = 0; i < calculatedFieldList.Count; i++) {
                        allFieldList.Add(calculatedFieldList[i].Name);
                    }

                    allFields = allFieldList;
                    headerIndex = BuildFieldIndex(allFieldList);
                }

                var dataFieldList = ToNonNullList(dataFields);
                if (dataFieldList.Count == 0) {
                    dataFieldList.Add(new ExcelPivotDataField(headers[headers.Count - 1], DataConsolidateFunctionValues.Sum));
                }
                var pivotFilterList = ToNonNullList(pivotFilters);
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

                if (generatedFieldsBySource.Count != 0) {
                    (rowFieldIndices, columnFieldIndices, pageFieldIndices) = ExpandGeneratedGroupingFieldIndices(
                        rowFieldIndices,
                        columnFieldIndices,
                        pageFieldIndices,
                        generatedFieldsBySource);
                }
                ReportPivotTiming("AddPivotTable.ResolveFields");

                var dataFieldIndices = new HashSet<int>();
                foreach (var df in dataFieldList) {
                    int idx = ResolveFieldIndex(df.FieldName, headerIndex, nameof(dataFields));
                    dataFieldIndices.Add(idx);
                }

                var fieldOptionMap = BuildPivotFieldOptionMap(fieldOptions, headerIndex);
                bool[] sourceSharedItemRequirements = BuildPivotSharedItemRequirements(
                    headers.Count,
                    rowFieldIndices,
                    columnFieldIndices,
                    pageFieldIndices,
                    groupingMap,
                    pivotFilterList,
                    headerIndex,
                    fieldOptionMap);

                var workbookPart = WorkbookPartRoot;
                var workbook = workbookPart.Workbook ??= new Workbook();
                _pivotStylesCache = null;
                bool canUseDeferredPivotValues = deferredPivotSource != null
                    && groupingMap.Count == 0
                    && generatedGroupingFields.Count == 0;
                if (deferredPivotSource != null && !canUseDeferredPivotValues) {
                    _excelDocument.MaterializeDeferredDataSetImport();
                }

                var fieldValueMap = canUseDeferredPivotValues
                    ? BuildPivotFieldValueMap(deferredPivotSource!, headers.Count, r1 + 1, r2, c1, sourceSharedItemRequirements)
                    : BuildPivotFieldValueMap(headers.Count, r1 + 1, r2, c1, groupingMap, sourceSharedItemRequirements);
                var generatedFieldValueMap = BuildGeneratedPivotFieldValueMap(generatedGroupingFields, r1 + 1, r2, c1);
                ReportPivotTiming("AddPivotTable.BuildFieldValueMap");
                if (deferredPivotSource != null && canUseDeferredPivotValues) {
                    _excelDocument.PreserveDeferredDataSetFastSaveModelAndClearCandidate();
                }
                ReportPivotTiming("AddPivotTable.PreserveFastSaveModel");

                var allFieldValueMap = BuildPivotTextValueMap(fieldValueMap, generatedFieldValueMap, calculatedFieldList.Count, allFields.Count);
                ExpandGeneratedGroupingFieldOptions(fieldOptionMap, generatedFieldsBySource, allFields, allFieldValueMap);
                uint cacheId = NextPivotCacheId(workbookPart);
                ReportPivotTiming("AddPivotTable.PrepareCacheMetadata");

                int sourceRecordCount = Math.Max(0, r2 - r1);
                bool effectiveSavePivotCacheRecords = options?.SaveSourceData == true;
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
                    RecordCount = (uint)sourceRecordCount,
                    SaveData = effectiveSavePivotCacheRecords,
                    RefreshOnLoad = options?.RefreshOnOpen ?? !effectiveSavePivotCacheRecords
                };

                for (int i = 0; i < headers.Count; i++) {
                    string header = headers[i];
                    var cacheField = new CacheField { Name = header };
                    groupingMap.TryGetValue(i, out var grouping);
                    cacheField.SharedItems = BuildSharedItems(fieldValueMap[i], grouping, sourceSharedItemRequirements[i]);
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
                ReportPivotTiming("AddPivotTable.BuildCacheFields");

                cacheDefPart.PivotCacheDefinition = cacheDef;
                ReportPivotTiming("AddPivotTable.SaveCacheDefinition");
                var cacheRecordsPart = cacheDefPart.AddNewPart<PivotTableCacheRecordsPart>();
                if (!effectiveSavePivotCacheRecords) {
                    cacheRecordsPart.PivotCacheRecords = new PivotCacheRecords { Count = 0U };
                } else if (canUseDeferredPivotValues) {
                    WritePivotCacheRecords(
                        cacheRecordsPart,
                        deferredPivotSource!,
                        headers.Count,
                        r1 + 1,
                        r2,
                        c1,
                        fieldValueMap,
                        sourceSharedItemRequirements,
                        generatedGroupingFields,
                        generatedFieldValueMap,
                        calculatedFieldList.Count);
                } else {
                    cacheRecordsPart.PivotCacheRecords = BuildPivotCacheRecords(
                        headers.Count,
                        r1 + 1,
                        r2,
                        c1,
                        groupingMap,
                        fieldValueMap,
                        sourceSharedItemRequirements,
                        generatedGroupingFields,
                        generatedFieldValueMap,
                        calculatedFieldList.Count);
                }
                ReportPivotTiming("AddPivotTable.BuildAndSaveCacheRecords");

                var pivotCaches = workbook.PivotCaches ?? workbook.AppendChild(new PivotCaches());
                pivotCaches.Append(new PivotCache {
                    CacheId = cacheId,
                    Id = workbookPart.GetIdOfPart(cacheDefPart)
                });
                // Count attribute is optional; OpenXml SDK does not expose a setter for PivotCaches.Count in all targets.

                string pivotName = EnsureUniquePivotTableName(name, _worksheetPart.PivotTableParts);

                var pivotPart = _worksheetPart.AddNewPart<PivotTablePart>();
                pivotPart.AddPart(cacheDefPart);

                var pivotFields = new PivotFields { Count = (uint)allFields.Count };
                for (int i = 0; i < allFields.Count; i++) {
                    ExcelPivotFieldOptions? options = null;
                    if (fieldOptionMap != null) {
                        fieldOptionMap.TryGetValue(i, out options);
                    }

                    var pivotField = new PivotField { ShowAll = options?.ShowAll ?? true };
                    if (pageFieldIndices.Contains(i)) pivotField.Axis = PivotTableAxisValues.AxisPage;
                    if (rowFieldIndices.Contains(i)) pivotField.Axis = PivotTableAxisValues.AxisRow;
                    if (columnFieldIndices.Contains(i)) pivotField.Axis = PivotTableAxisValues.AxisColumn;
                    if (dataFieldIndices.Contains(i)) pivotField.DataField = true;
                    IReadOnlyList<string> values = options != null ? allFieldValueMap[i] : Array.Empty<string>();
                    ApplyPivotFieldOptions(pivotField, options, workbookPart, values);
                    pivotFields.Append(pivotField);
                }
                ReportPivotTiming("AddPivotTable.BuildPivotFields");

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
                        ExcelPivotFieldOptions? options = null;
                        if (fieldOptionMap != null) {
                            fieldOptionMap.TryGetValue(idx, out options);
                        }

                        IReadOnlyList<string> values = options?.SelectedItem != null ? allFieldValueMap[idx] : Array.Empty<string>();
                        pageFieldsElement.Append(CreatePageField(idx, options, values));
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

                PivotFilters? pivotFiltersElement = CreatePivotFilters(pivotFilterList, headerIndex, dataFieldList, _excelDocument.DateSystem);

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
                    PreserveFormatting = options?.PreserveFormatting ?? true,
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
                if (options?.EnableDrill.HasValue == true) pivotDefinition.EnableDrill = options.EnableDrill.Value;

                pivotPart.PivotTableDefinition = pivotDefinition;
                ReportPivotTiming("AddPivotTable.BuildAndSavePivotDefinition");
            });
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
    }
}
