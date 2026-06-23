using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel.Fluent {
    /// <summary>
    /// Fluent builder for creating pivot tables from an A1 source range.
    /// </summary>
    public sealed class PivotTableBuilder {
        private readonly ExcelSheet _sheet;
        private readonly string _sourceRange;
        private readonly List<string> _rowFields = new();
        private readonly List<string> _columnFields = new();
        private readonly List<string> _pageFields = new();
        private readonly List<ExcelPivotDataField> _dataFields = new();
        private readonly List<ExcelPivotFieldOptions> _fieldOptions = new();
        private readonly List<ExcelPivotFilter> _pivotFilters = new();
        private readonly List<ExcelPivotCalculatedField> _calculatedFields = new();
        private readonly List<ExcelPivotGrouping> _groupings = new();
        private bool _showRowGrandTotals = true;
        private bool _showColumnGrandTotals = true;
        private string? _pivotStyleName;
        private ExcelPivotLayout _layout = ExcelPivotLayout.Compact;
        private bool? _dataOnRows;
        private bool? _showHeaders;
        private bool? _showEmptyRows;
        private bool? _showEmptyColumns;
        private bool? _showDrill;
        private string? _rowHeaderCaption;
        private string? _columnHeaderCaption;
        private string? _grandTotalCaption;
        private string? _missingCaption;
        private string? _errorCaption;
        private bool? _showDataDropDown;
        private bool? _showDropZones;
        private bool? _showDataTips;
        private bool? _showMemberPropertyTips;
        private bool? _fieldListSortAscending;
        private bool? _customListSort;
        private bool? _refreshOnOpen;
        private bool? _saveSourceData;
        private bool? _preserveFormatting;
        private bool? _enableDrill;

        internal PivotTableBuilder(ExcelSheet sheet, string sourceRange) {
            _sheet = sheet ?? throw new ArgumentNullException(nameof(sheet));
            _sourceRange = string.IsNullOrWhiteSpace(sourceRange)
                ? throw new ArgumentNullException(nameof(sourceRange))
                : sourceRange;
        }

        /// <summary>Adds one or more row fields by source header name.</summary>
        public PivotTableBuilder Rows(params string[] fieldNames) {
            AddFieldNames(_rowFields, fieldNames, nameof(fieldNames));
            return this;
        }

        /// <summary>Adds one or more column fields by source header name.</summary>
        public PivotTableBuilder Columns(params string[] fieldNames) {
            AddFieldNames(_columnFields, fieldNames, nameof(fieldNames));
            return this;
        }

        /// <summary>Adds one or more page/filter fields by source header name.</summary>
        public PivotTableBuilder Filters(params string[] fieldNames) {
            AddFieldNames(_pageFields, fieldNames, nameof(fieldNames));
            return this;
        }

        /// <summary>Adds one or more page/filter fields by source header name.</summary>
        public PivotTableBuilder Pages(params string[] fieldNames) => Filters(fieldNames);

        /// <summary>Adds a Sum data field.</summary>
        public PivotTableBuilder Sum(string fieldName, string? displayName = null, string? numberFormat = null) {
            return Value(fieldName, DataConsolidateFunctionValues.Sum, displayName, numberFormat: numberFormat);
        }

        /// <summary>Adds a Count data field.</summary>
        public PivotTableBuilder Count(string fieldName, string? displayName = null, string? numberFormat = null) {
            return Value(fieldName, DataConsolidateFunctionValues.Count, displayName, numberFormat: numberFormat);
        }

        /// <summary>Adds an Average data field.</summary>
        public PivotTableBuilder Average(string fieldName, string? displayName = null, string? numberFormat = null) {
            return Value(fieldName, DataConsolidateFunctionValues.Average, displayName, numberFormat: numberFormat);
        }

        /// <summary>Adds a Min data field.</summary>
        public PivotTableBuilder Min(string fieldName, string? displayName = null, string? numberFormat = null) {
            return Value(fieldName, DataConsolidateFunctionValues.Minimum, displayName, numberFormat: numberFormat);
        }

        /// <summary>Adds a Max data field.</summary>
        public PivotTableBuilder Max(string fieldName, string? displayName = null, string? numberFormat = null) {
            return Value(fieldName, DataConsolidateFunctionValues.Maximum, displayName, numberFormat: numberFormat);
        }

        /// <summary>Adds a data field using a specific aggregation function.</summary>
        public PivotTableBuilder Value(
            string fieldName,
            DataConsolidateFunctionValues function,
            string? displayName = null,
            uint? numberFormatId = null,
            string? numberFormat = null,
            ShowDataAsValues? showDataAs = null,
            int? baseField = null,
            uint? baseItem = null) {
            if (string.IsNullOrWhiteSpace(fieldName)) throw new ArgumentNullException(nameof(fieldName));
            _dataFields.Add(new ExcelPivotDataField(fieldName, function, displayName, numberFormatId, numberFormat, showDataAs, baseField, baseItem));
            return this;
        }

        /// <summary>Adds a Sum data field shown as a percentage of the pivot grand total.</summary>
        public PivotTableBuilder PercentOfTotal(string fieldName, string? displayName = null, string? numberFormat = "0.0%") {
            return Value(fieldName, DataConsolidateFunctionValues.Sum, displayName, numberFormat: numberFormat, showDataAs: ShowDataAsValues.PercentOfTotal);
        }

        /// <summary>Sets the pivot table style name, for example PivotStyleMedium9.</summary>
        public PivotTableBuilder Style(string pivotStyleName) {
            _pivotStyleName = string.IsNullOrWhiteSpace(pivotStyleName)
                ? throw new ArgumentNullException(nameof(pivotStyleName))
                : pivotStyleName;
            return this;
        }

        /// <summary>Sets the pivot table layout.</summary>
        public PivotTableBuilder Layout(ExcelPivotLayout layout) {
            _layout = layout;
            return this;
        }

        /// <summary>Controls row and column grand totals.</summary>
        public PivotTableBuilder GrandTotals(bool rows = true, bool columns = true) {
            _showRowGrandTotals = rows;
            _showColumnGrandTotals = columns;
            return this;
        }

        /// <summary>Controls common pivot display flags.</summary>
        public PivotTableBuilder Display(
            bool? dataOnRows = null,
            bool? showHeaders = null,
            bool? showEmptyRows = null,
            bool? showEmptyColumns = null,
            bool? showDrill = null,
            bool? showDataDropDown = null,
            bool? showDropZones = null,
            bool? showDataTips = null,
            bool? showMemberPropertyTips = null,
            bool? fieldListSortAscending = null,
            bool? customListSort = null) {
            _dataOnRows = dataOnRows;
            _showHeaders = showHeaders;
            _showEmptyRows = showEmptyRows;
            _showEmptyColumns = showEmptyColumns;
            _showDrill = showDrill;
            _showDataDropDown = showDataDropDown;
            _showDropZones = showDropZones;
            _showDataTips = showDataTips;
            _showMemberPropertyTips = showMemberPropertyTips;
            _fieldListSortAscending = fieldListSortAscending;
            _customListSort = customListSort;
            return this;
        }

        /// <summary>Controls pivot cache refresh and interactive workbook behavior.</summary>
        public PivotTableBuilder Interaction(
            bool? refreshOnOpen = null,
            bool? saveSourceData = null,
            bool? preserveFormatting = null,
            bool? enableDrill = null) {
            _refreshOnOpen = refreshOnOpen;
            _saveSourceData = saveSourceData;
            _preserveFormatting = preserveFormatting;
            _enableDrill = enableDrill;
            return this;
        }

        /// <summary>Sets pivot table captions.</summary>
        public PivotTableBuilder Captions(
            string? rowHeader = null,
            string? columnHeader = null,
            string? grandTotal = null,
            string? missing = null,
            string? error = null) {
            _rowHeaderCaption = rowHeader;
            _columnHeaderCaption = columnHeader;
            _grandTotalCaption = grandTotal;
            _missingCaption = missing;
            _errorCaption = error;
            return this;
        }

        /// <summary>Adds field-level display, sort, format, and item-filter options.</summary>
        public PivotTableBuilder FieldOptions(params ExcelPivotFieldOptions[] options) {
            if (options == null) throw new ArgumentNullException(nameof(options));
            foreach (var option in options) {
                if (option != null) {
                    _fieldOptions.Add(option);
                }
            }

            return this;
        }

        /// <summary>Hides specific source items for a pivot field.</summary>
        public PivotTableBuilder HideItems(string fieldName, params string[] items) {
            return UpdateFieldOptions(
                fieldName,
                hiddenItems: items,
                replaceHiddenItems: true,
                visibleItems: Array.Empty<string>(),
                replaceVisibleItems: true);
        }

        /// <summary>Shows only specific source items for a pivot field.</summary>
        public PivotTableBuilder ShowOnlyItems(string fieldName, params string[] items) {
            return UpdateFieldOptions(
                fieldName,
                hiddenItems: Array.Empty<string>(),
                replaceHiddenItems: true,
                visibleItems: items,
                replaceVisibleItems: true);
        }

        /// <summary>Selects an item for a page/filter field and adds the field to the page area if needed.</summary>
        public PivotTableBuilder SelectPageItem(string fieldName, string item) {
            if (string.IsNullOrWhiteSpace(fieldName)) throw new ArgumentNullException(nameof(fieldName));
            if (string.IsNullOrWhiteSpace(item)) throw new ArgumentNullException(nameof(item));
            if (!_pageFields.Contains(fieldName, StringComparer.OrdinalIgnoreCase)) {
                _pageFields.Add(fieldName);
            }

            return UpdateFieldOptions(fieldName, selectedItem: item, replaceSelectedItem: true);
        }

        /// <summary>Sets the sort mode for a pivot field.</summary>
        public PivotTableBuilder SortField(string fieldName, FieldSortValues sortType) {
            return UpdateFieldOptions(fieldName, sortType: sortType);
        }

        /// <summary>Controls whether the pivot field uses its default subtotal.</summary>
        public PivotTableBuilder Subtotals(string fieldName, bool enabled = true) {
            return UpdateFieldOptions(fieldName, defaultSubtotal: enabled);
        }

        /// <summary>Controls whether subtotals are shown at the top for a pivot field.</summary>
        public PivotTableBuilder SubtotalsAtTop(string fieldName, bool enabled = true) {
            return UpdateFieldOptions(fieldName, subtotalTop: enabled);
        }

        /// <summary>Controls compact and outline layout flags for a pivot field.</summary>
        public PivotTableBuilder FieldLayout(string fieldName, bool? compact = null, bool? outline = null) {
            return UpdateFieldOptions(fieldName, compact: compact, outline: outline);
        }

        /// <summary>Controls blank-row and page-break insertion after pivot field items.</summary>
        public PivotTableBuilder FieldBreaks(string fieldName, bool? insertBlankRow = null, bool? insertPageBreak = null) {
            return UpdateFieldOptions(fieldName, insertBlankRow: insertBlankRow, insertPageBreak: insertPageBreak);
        }

        /// <summary>Controls common pivot field display flags.</summary>
        public PivotTableBuilder FieldDisplay(
            string fieldName,
            bool? showAll = null,
            bool? showDropDowns = null,
            bool? multipleItemSelectionAllowed = null,
            bool? includeNewItemsInFilter = null) {
            return UpdateFieldOptions(
                fieldName,
                showAll: showAll,
                showDropDowns: showDropDowns,
                multipleItemSelectionAllowed: multipleItemSelectionAllowed,
                includeNewItemsInFilter: includeNewItemsInFilter);
        }

        /// <summary>Sets a number format code for a pivot field.</summary>
        public PivotTableBuilder FieldNumberFormat(string fieldName, string numberFormat) {
            if (string.IsNullOrWhiteSpace(numberFormat)) throw new ArgumentNullException(nameof(numberFormat));
            return UpdateFieldOptions(fieldName, numberFormat: numberFormat, replaceNumberFormat: true);
        }

        /// <summary>Sets a built-in or custom number format id for a pivot field.</summary>
        public PivotTableBuilder FieldNumberFormatId(string fieldName, uint numberFormatId) {
            return UpdateFieldOptions(fieldName, numberFormatId: numberFormatId, replaceNumberFormat: true);
        }

        /// <summary>Sets a custom subtotal caption for a pivot field.</summary>
        public PivotTableBuilder SubtotalCaption(string fieldName, string caption) {
            if (string.IsNullOrWhiteSpace(caption)) throw new ArgumentNullException(nameof(caption));
            return UpdateFieldOptions(fieldName, subtotalCaption: caption, replaceSubtotalCaption: true);
        }

        /// <summary>Adds label or value filters to the pivot table.</summary>
        public PivotTableBuilder Filter(params ExcelPivotFilter[] filters) {
            if (filters == null) throw new ArgumentNullException(nameof(filters));
            foreach (var filter in filters) {
                if (filter != null) {
                    _pivotFilters.Add(filter);
                }
            }

            return this;
        }

        /// <summary>Adds a formula-backed calculated pivot field.</summary>
        public PivotTableBuilder CalculatedField(string name, string formula, string? caption = null, uint? numberFormatId = null, string? numberFormat = null) {
            _calculatedFields.Add(new ExcelPivotCalculatedField(name, formula, caption, numberFormatId, numberFormat));
            return this;
        }

        /// <summary>Adds date or numeric grouping metadata for pivot fields.</summary>
        public PivotTableBuilder Group(params ExcelPivotGrouping[] groupings) {
            if (groupings == null) throw new ArgumentNullException(nameof(groupings));
            foreach (var grouping in groupings) {
                if (grouping != null) {
                    _groupings.Add(grouping);
                }
            }

            return this;
        }

        /// <summary>Adds date grouping metadata for a pivot field.</summary>
        public PivotTableBuilder DateGroup(string fieldName, GroupByValues groupBy, DateTime? startDate = null, DateTime? endDate = null, double? interval = null) {
            _groupings.Add(ExcelPivotGrouping.Date(fieldName, groupBy, startDate, endDate, interval));
            return this;
        }

        /// <summary>Adds generated date hierarchy fields for a pivot field, such as years, quarters, and months.</summary>
        public PivotTableBuilder DateHierarchy(string fieldName, params GroupByValues[] levels) {
            _groupings.Add(ExcelPivotGrouping.DateHierarchy(fieldName, levels));
            return this;
        }

        /// <summary>Adds numeric range grouping metadata for a pivot field.</summary>
        public PivotTableBuilder NumberGroup(string fieldName, double interval, double? startNumber = null, double? endNumber = null) {
            _groupings.Add(ExcelPivotGrouping.Number(fieldName, interval, startNumber, endNumber));
            return this;
        }

        /// <summary>Creates the pivot table at the destination cell and returns the source sheet.</summary>
        public ExcelSheet At(string destinationCell, string? name = null) {
            if (string.IsNullOrWhiteSpace(destinationCell)) throw new ArgumentNullException(nameof(destinationCell));

            _sheet.AddPivotTable(
                sourceRange: _sourceRange,
                destinationCell: destinationCell,
                name: name,
                rowFields: _rowFields.Count == 0 ? null : _rowFields,
                columnFields: _columnFields.Count == 0 ? null : _columnFields,
                pageFields: _pageFields.Count == 0 ? null : _pageFields,
                dataFields: _dataFields.Count == 0 ? null : _dataFields,
                showRowGrandTotals: _showRowGrandTotals,
                showColumnGrandTotals: _showColumnGrandTotals,
                pivotStyleName: _pivotStyleName,
                layout: _layout,
                dataOnRows: _dataOnRows,
                showHeaders: _showHeaders,
                showEmptyRows: _showEmptyRows,
                showEmptyColumns: _showEmptyColumns,
                showDrill: _showDrill,
                fieldOptions: _fieldOptions.Count == 0 ? null : _fieldOptions,
                rowHeaderCaption: _rowHeaderCaption,
                columnHeaderCaption: _columnHeaderCaption,
                grandTotalCaption: _grandTotalCaption,
                missingCaption: _missingCaption,
                errorCaption: _errorCaption,
                showDataDropDown: _showDataDropDown,
                showDropZones: _showDropZones,
                showDataTips: _showDataTips,
                showMemberPropertyTips: _showMemberPropertyTips,
                fieldListSortAscending: _fieldListSortAscending,
                customListSort: _customListSort,
                pivotFilters: _pivotFilters.Count == 0 ? null : _pivotFilters,
                calculatedFields: _calculatedFields.Count == 0 ? null : _calculatedFields,
                groupings: _groupings.Count == 0 ? null : _groupings,
                options: CreateOptions());

            return _sheet;
        }

        private ExcelPivotTableOptions? CreateOptions() {
            if (!_refreshOnOpen.HasValue
                && !_saveSourceData.HasValue
                && !_preserveFormatting.HasValue
                && !_enableDrill.HasValue) {
                return null;
            }

            return new ExcelPivotTableOptions {
                RefreshOnOpen = _refreshOnOpen,
                SaveSourceData = _saveSourceData,
                PreserveFormatting = _preserveFormatting,
                EnableDrill = _enableDrill
            };
        }

        private static void AddFieldNames(List<string> target, string[] fieldNames, string parameterName) {
            if (fieldNames == null) throw new ArgumentNullException(parameterName);
            foreach (string fieldName in fieldNames) {
                if (string.IsNullOrWhiteSpace(fieldName)) {
                    throw new ArgumentException("Pivot field names cannot be empty.", parameterName);
                }

                target.Add(fieldName);
            }
        }

        private PivotTableBuilder UpdateFieldOptions(
            string fieldName,
            FieldSortValues? sortType = null,
            uint? numberFormatId = null,
            string? numberFormat = null,
            bool replaceNumberFormat = false,
            bool? showAll = null,
            bool? defaultSubtotal = null,
            bool? subtotalTop = null,
            bool? insertBlankRow = null,
            bool? insertPageBreak = null,
            bool? compact = null,
            bool? outline = null,
            bool? showDropDowns = null,
            bool? multipleItemSelectionAllowed = null,
            bool? includeNewItemsInFilter = null,
            string? subtotalCaption = null,
            bool replaceSubtotalCaption = false,
            IEnumerable<string>? hiddenItems = null,
            bool replaceHiddenItems = false,
            IEnumerable<string>? visibleItems = null,
            bool replaceVisibleItems = false,
            string? selectedItem = null,
            bool replaceSelectedItem = false) {
            if (string.IsNullOrWhiteSpace(fieldName)) throw new ArgumentNullException(nameof(fieldName));

            int existingIndex = -1;
            for (int i = _fieldOptions.Count - 1; i >= 0; i--) {
                if (string.Equals(_fieldOptions[i].FieldName, fieldName, StringComparison.OrdinalIgnoreCase)) {
                    existingIndex = i;
                    break;
                }
            }

            ExcelPivotFieldOptions? existing = existingIndex >= 0 ? _fieldOptions[existingIndex] : null;
            var updated = new ExcelPivotFieldOptions(
                fieldName,
                sortType: sortType ?? existing?.SortType,
                numberFormatId: replaceNumberFormat ? numberFormatId : existing?.NumberFormatId,
                numberFormat: replaceNumberFormat ? numberFormat : existing?.NumberFormat,
                showAll: showAll ?? existing?.ShowAll,
                defaultSubtotal: defaultSubtotal ?? existing?.DefaultSubtotal,
                subtotalTop: subtotalTop ?? existing?.SubtotalTop,
                insertBlankRow: insertBlankRow ?? existing?.InsertBlankRow,
                insertPageBreak: insertPageBreak ?? existing?.InsertPageBreak,
                compact: compact ?? existing?.Compact,
                outline: outline ?? existing?.Outline,
                showDropDowns: showDropDowns ?? existing?.ShowDropDowns,
                multipleItemSelectionAllowed: multipleItemSelectionAllowed ?? existing?.MultipleItemSelectionAllowed,
                includeNewItemsInFilter: includeNewItemsInFilter ?? existing?.IncludeNewItemsInFilter,
                subtotalCaption: replaceSubtotalCaption ? subtotalCaption : existing?.SubtotalCaption,
                hiddenItems: replaceHiddenItems ? hiddenItems : existing?.HiddenItems,
                visibleItems: replaceVisibleItems ? visibleItems : existing?.VisibleItems,
                selectedItem: replaceSelectedItem ? selectedItem : existing?.SelectedItem);

            if (existingIndex >= 0) {
                _fieldOptions[existingIndex] = updated;
            } else {
                _fieldOptions.Add(updated);
            }

            return this;
        }
    }
}
