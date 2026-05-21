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
            string? numberFormat = null) {
            if (string.IsNullOrWhiteSpace(fieldName)) throw new ArgumentNullException(nameof(fieldName));
            _dataFields.Add(new ExcelPivotDataField(fieldName, function, displayName, numberFormatId, numberFormat));
            return this;
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
                customListSort: _customListSort);

            return _sheet;
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
    }
}
