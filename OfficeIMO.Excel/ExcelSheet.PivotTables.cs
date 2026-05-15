using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        /// <summary>
        /// Returns pivot tables defined on this worksheet.
        /// </summary>
        public IReadOnlyList<ExcelPivotTableInfo> GetPivotTables() {
            return Locking.ExecuteRead(_excelDocument.EnsureLock(), () => {
                var list = new List<ExcelPivotTableInfo>();
                var workbookPart = WorkbookPartRoot;

                var cacheMap = BuildPivotCacheMap(workbookPart);
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
                    var dataFields = ResolveDataFields(def.DataFields?.Elements<DataField>(), cacheFields);
                    var fieldInfos = ResolveFieldInfos(def.PivotFields?.Elements<PivotField>(), cacheFields, BuildCacheFieldItems(cacheDef));

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
                        fields: fieldInfos));
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
                customListSort: null);
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
            bool? customListSort = null) {
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

                var headerIndex = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
                for (int i = 0; i < headers.Count; i++) {
                    headerIndex[headers[i]] = i;
                }

                var dataFieldList = (dataFields ?? Array.Empty<ExcelPivotDataField>()).Where(df => df != null).ToList();
                if (dataFieldList.Count == 0) {
                    dataFieldList.Add(new ExcelPivotDataField(headers[headers.Count - 1], DataConsolidateFunctionValues.Sum));
                }

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

                var dataFieldIndices = new HashSet<int>();
                foreach (var df in dataFieldList) {
                    int idx = ResolveFieldIndex(df.FieldName, headerIndex, nameof(dataFields));
                    dataFieldIndices.Add(idx);
                }

                var workbookPart = WorkbookPartRoot;
                var workbook = workbookPart.Workbook ??= new Workbook();
                var fieldOptionMap = BuildPivotFieldOptionMap(fieldOptions, headerIndex);
                var fieldValueMap = BuildPivotFieldValueMap(headers.Count, r1 + 1, r2, c1);
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
                    CacheFields = new CacheFields { Count = (uint)headers.Count },
                    RecordCount = 0,
                    RefreshOnLoad = true,
                    SaveData = false
                };

                for (int i = 0; i < headers.Count; i++) {
                    string header = headers[i];
                    var cacheField = new CacheField { Name = header };
                    cacheField.SharedItems = BuildSharedItems(fieldValueMap[i]);
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

                var pivotFields = new PivotFields { Count = (uint)headers.Count };
                for (int i = 0; i < headers.Count; i++) {
                    fieldOptionMap.TryGetValue(i, out var options);
                    var pivotField = new PivotField { ShowAll = options?.ShowAll ?? true };
                    if (pageFieldIndices.Contains(i)) pivotField.Axis = PivotTableAxisValues.AxisPage;
                    if (rowFieldIndices.Contains(i)) pivotField.Axis = PivotTableAxisValues.AxisRow;
                    if (columnFieldIndices.Contains(i)) pivotField.Axis = PivotTableAxisValues.AxisColumn;
                    if (dataFieldIndices.Contains(i)) pivotField.DataField = true;
                    ApplyPivotFieldOptions(pivotField, options, workbookPart, fieldValueMap[i]);
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
                        pageFieldsElement.Append(CreatePageField(idx, options, fieldValueMap[idx]));
                    }
                }

                var dataFieldsElement = new DataFields { Count = (uint)dataFieldList.Count };
                foreach (var df in dataFieldList) {
                    int idx = ResolveFieldIndex(df.FieldName, headerIndex, nameof(dataFields));
                    string display = df.DisplayName ?? $"{df.Function} of {headers[idx]}";
                    var dataField = new DataField {
                        Name = display,
                        Field = (uint)idx,
                        Subtotal = df.Function
                    };
                    uint? numberFormatId = ResolveNumberFormatId(workbookPart, df.NumberFormatId, df.NumberFormat);
                    if (numberFormatId.HasValue) dataField.NumberFormatId = numberFormatId.Value;
                    dataFieldsElement.Append(dataField);
                }

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

        private List<IReadOnlyList<string>> BuildPivotFieldValueMap(int fieldCount, int firstDataRow, int lastDataRow, int firstColumn) {
            var maps = new List<IReadOnlyList<string>>(fieldCount);
            for (int field = 0; field < fieldCount; field++) {
                var values = new List<string>();
                var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                int column = firstColumn + field;
                for (int row = firstDataRow; row <= lastDataRow; row++) {
                    string text = TryGetCellText(row, column, out string cellText) ? cellText.Trim() : string.Empty;
                    if (seen.Add(text)) {
                        values.Add(text);
                    }
                }

                maps.Add(values);
            }

            return maps;
        }

        private static SharedItems BuildSharedItems(IReadOnlyList<string> values) {
            bool hasBlank = values.Any(string.IsNullOrEmpty);
            var sharedItems = new SharedItems {
                Count = (uint)values.Count,
                ContainsString = values.Any(value => !string.IsNullOrEmpty(value)),
                ContainsBlank = hasBlank
            };

            foreach (string value in values) {
                if (string.IsNullOrEmpty(value)) {
                    sharedItems.Append(new MissingItem());
                } else {
                    sharedItems.Append(new StringItem { Val = value });
                }
            }

            return sharedItems;
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

        private static List<ExcelPivotFieldInfo> ResolveFieldInfos(IEnumerable<PivotField>? fields, IReadOnlyList<string> cacheFields,
            IReadOnlyList<IReadOnlyList<string>> cacheFieldItems) {
            var list = new List<ExcelPivotFieldInfo>();
            if (fields == null) return list;
            int index = 0;
            foreach (var field in fields) {
                IReadOnlyList<string> itemValues = index < cacheFieldItems.Count ? cacheFieldItems[index] : Array.Empty<string>();
                list.Add(new ExcelPivotFieldInfo(
                    fieldName: ResolveFieldName(index, cacheFields),
                    axis: field.Axis?.Value,
                    sortType: field.SortType?.Value,
                    numberFormatId: field.NumberFormatId?.Value,
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
                    hiddenItems: ResolveHiddenItems(field.Items, itemValues)));
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

        private static List<ExcelPivotDataFieldInfo> ResolveDataFields(IEnumerable<DataField>? fields, IReadOnlyList<string> cacheFields) {
            var list = new List<ExcelPivotDataFieldInfo>();
            if (fields == null) return list;
            foreach (var field in fields) {
                int idx = field.Field?.Value is uint u ? (int)u : 0;
                var name = ResolveFieldName(idx, cacheFields);
                var fn = field.Subtotal?.Value ?? DataConsolidateFunctionValues.Sum;
                var display = field.Name?.Value;
                list.Add(new ExcelPivotDataFieldInfo(name, fn, display, field.NumberFormatId?.Value));
            }
            return list;
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
