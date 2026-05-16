using System.Collections.Generic;

namespace OfficeIMO.Excel {
    /// <summary>
    /// Represents an existing pivot table definition.
    /// </summary>
    public sealed class ExcelPivotTableInfo {
        /// <summary>
        /// Creates a pivot table info instance.
        /// </summary>
        public ExcelPivotTableInfo(string name,
            uint cacheId,
            string? location,
            string? sourceSheet,
            string? sourceRange,
            string sheetName,
            int sheetIndex,
            string? pivotStyle,
            ExcelPivotLayout layout,
            bool? dataOnRows,
            bool? showHeaders,
            bool? showEmptyRows,
            bool? showEmptyColumns,
            bool? showDrill,
            IReadOnlyList<string> rowFields,
            IReadOnlyList<string> columnFields,
            IReadOnlyList<string> pageFields,
            IReadOnlyList<ExcelPivotDataFieldInfo> dataFields)
            : this(name, cacheId, location, sourceSheet, sourceRange, sheetName, sheetIndex, pivotStyle,
                layout, dataOnRows, showHeaders, showEmptyRows, showEmptyColumns, showDrill,
                null, null, null, null, null, null, null, null, null, null, null, null, null,
                rowFields, columnFields, pageFields, dataFields, null) {
        }

        /// <summary>
        /// Creates a pivot table info instance.
        /// </summary>
        public ExcelPivotTableInfo(string name,
            uint cacheId,
            string? location,
            string? sourceSheet,
            string? sourceRange,
            string sheetName,
            int sheetIndex,
            string? pivotStyle,
            ExcelPivotLayout layout,
            bool? dataOnRows,
            bool? showHeaders,
            bool? showEmptyRows,
            bool? showEmptyColumns,
            bool? showDrill,
            bool? rowGrandTotals,
            bool? columnGrandTotals,
            string? rowHeaderCaption,
            string? columnHeaderCaption,
            string? grandTotalCaption,
            string? missingCaption,
            string? errorCaption,
            bool? showDataDropDown,
            bool? showDropZones,
            bool? showDataTips,
            bool? showMemberPropertyTips,
            bool? fieldListSortAscending,
            bool? customListSort,
            IReadOnlyList<string> rowFields,
            IReadOnlyList<string> columnFields,
            IReadOnlyList<string> pageFields,
            IReadOnlyList<ExcelPivotDataFieldInfo> dataFields,
            IReadOnlyList<ExcelPivotFieldInfo>? fields = null) {
            Name = name;
            CacheId = cacheId;
            Location = location;
            SourceSheet = sourceSheet;
            SourceRange = sourceRange;
            SheetName = sheetName;
            SheetIndex = sheetIndex;
            PivotStyle = pivotStyle;
            Layout = layout;
            DataOnRows = dataOnRows;
            ShowHeaders = showHeaders;
            ShowEmptyRows = showEmptyRows;
            ShowEmptyColumns = showEmptyColumns;
            ShowDrill = showDrill;
            RowGrandTotals = rowGrandTotals;
            ColumnGrandTotals = columnGrandTotals;
            RowHeaderCaption = rowHeaderCaption;
            ColumnHeaderCaption = columnHeaderCaption;
            GrandTotalCaption = grandTotalCaption;
            MissingCaption = missingCaption;
            ErrorCaption = errorCaption;
            ShowDataDropDown = showDataDropDown;
            ShowDropZones = showDropZones;
            ShowDataTips = showDataTips;
            ShowMemberPropertyTips = showMemberPropertyTips;
            FieldListSortAscending = fieldListSortAscending;
            CustomListSort = customListSort;
            RowFields = rowFields;
            ColumnFields = columnFields;
            PageFields = pageFields;
            DataFields = dataFields;
            Fields = fields ?? Array.Empty<ExcelPivotFieldInfo>();
        }

        /// <summary>
        /// Gets the pivot table name.
        /// </summary>
        public string Name { get; }

        /// <summary>
        /// Gets the cache id for the pivot table.
        /// </summary>
        public uint CacheId { get; }

        /// <summary>
        /// Gets the pivot table location (A1 range).
        /// </summary>
        public string? Location { get; }

        /// <summary>
        /// Gets the source sheet name from the cache.
        /// </summary>
        public string? SourceSheet { get; }

        /// <summary>
        /// Gets the source range from the cache.
        /// </summary>
        public string? SourceRange { get; }

        /// <summary>
        /// Gets the sheet name where the pivot table is defined.
        /// </summary>
        public string SheetName { get; }

        /// <summary>
        /// Gets the 0-based sheet index where the pivot table is defined.
        /// </summary>
        public int SheetIndex { get; }

        /// <summary>
        /// Gets the pivot table style name.
        /// </summary>
        public string? PivotStyle { get; }

        /// <summary>
        /// Gets the pivot layout mode.
        /// </summary>
        public ExcelPivotLayout Layout { get; }

        /// <summary>
        /// Gets whether data fields are shown on rows.
        /// </summary>
        public bool? DataOnRows { get; }

        /// <summary>
        /// Gets whether field headers are shown.
        /// </summary>
        public bool? ShowHeaders { get; }

        /// <summary>
        /// Gets whether empty rows are shown.
        /// </summary>
        public bool? ShowEmptyRows { get; }

        /// <summary>
        /// Gets whether empty columns are shown.
        /// </summary>
        public bool? ShowEmptyColumns { get; }

        /// <summary>
        /// Gets whether drill indicators are shown.
        /// </summary>
        public bool? ShowDrill { get; }

        /// <summary>
        /// Gets whether row grand totals are shown.
        /// </summary>
        public bool? RowGrandTotals { get; }

        /// <summary>
        /// Gets whether column grand totals are shown.
        /// </summary>
        public bool? ColumnGrandTotals { get; }

        /// <summary>
        /// Gets the row header caption.
        /// </summary>
        public string? RowHeaderCaption { get; }

        /// <summary>
        /// Gets the column header caption.
        /// </summary>
        public string? ColumnHeaderCaption { get; }

        /// <summary>
        /// Gets the grand total caption.
        /// </summary>
        public string? GrandTotalCaption { get; }

        /// <summary>
        /// Gets the missing-value caption.
        /// </summary>
        public string? MissingCaption { get; }

        /// <summary>
        /// Gets the error-value caption.
        /// </summary>
        public string? ErrorCaption { get; }

        /// <summary>
        /// Gets whether the data drop-down is shown.
        /// </summary>
        public bool? ShowDataDropDown { get; }

        /// <summary>
        /// Gets whether drop zones are shown.
        /// </summary>
        public bool? ShowDropZones { get; }

        /// <summary>
        /// Gets whether data tips are shown.
        /// </summary>
        public bool? ShowDataTips { get; }

        /// <summary>
        /// Gets whether member-property tips are shown.
        /// </summary>
        public bool? ShowMemberPropertyTips { get; }

        /// <summary>
        /// Gets whether the field list sorts ascending.
        /// </summary>
        public bool? FieldListSortAscending { get; }

        /// <summary>
        /// Gets whether custom-list sorting is enabled.
        /// </summary>
        public bool? CustomListSort { get; }

        /// <summary>
        /// Gets row field names.
        /// </summary>
        public IReadOnlyList<string> RowFields { get; }

        /// <summary>
        /// Gets column field names.
        /// </summary>
        public IReadOnlyList<string> ColumnFields { get; }

        /// <summary>
        /// Gets page field names.
        /// </summary>
        public IReadOnlyList<string> PageFields { get; }

        /// <summary>
        /// Gets data fields.
        /// </summary>
        public IReadOnlyList<ExcelPivotDataFieldInfo> DataFields { get; }

        /// <summary>
        /// Gets detailed field metadata.
        /// </summary>
        public IReadOnlyList<ExcelPivotFieldInfo> Fields { get; }
    }
}
