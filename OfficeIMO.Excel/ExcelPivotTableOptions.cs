namespace OfficeIMO.Excel {
    /// <summary>
    /// Layout and workbook-interaction options for authored pivot tables.
    /// </summary>
    public sealed class ExcelPivotTableOptions {
        /// <summary>
        /// Gets or sets the zero-based placement of Values among fields on its row or column axis.
        /// Requires multiple measures. When omitted, Values is placed last.
        /// </summary>
        public int? ValuesAxisPosition { get; set; }

        /// <summary>
        /// Gets or sets whether Excel should refresh the pivot cache when the workbook opens.
        /// </summary>
        public bool? RefreshOnOpen { get; set; }

        /// <summary>
        /// Gets or sets whether source cache records should be saved in the workbook package. Defaults to false so
        /// source rows are not duplicated into the pivot cache unless the caller explicitly opts in.
        /// </summary>
        public bool? SaveSourceData { get; set; }

        /// <summary>
        /// Gets or sets whether Excel should preserve pivot table formatting during refreshes.
        /// </summary>
        public bool? PreserveFormatting { get; set; }

        /// <summary>
        /// Gets or sets whether drill-down interaction is enabled for pivot details.
        /// </summary>
        public bool? EnableDrill { get; set; }

        /// <summary>Gets or sets name for the authored pivot table.</summary>
        public string? Name { get; set; }

        /// <summary>Gets the rowFields used when creating the pivot table.</summary>
        public IList<string> RowFields { get; set; } = new List<string>();

        /// <summary>Gets the columnFields used when creating the pivot table.</summary>
        public IList<string> ColumnFields { get; set; } = new List<string>();

        /// <summary>Gets the pageFields used when creating the pivot table.</summary>
        public IList<string> PageFields { get; set; } = new List<string>();

        /// <summary>Gets the dataFields used when creating the pivot table.</summary>
        public IList<ExcelPivotDataField> DataFields { get; set; } = new List<ExcelPivotDataField>();

        /// <summary>Gets or sets showRowGrandTotals for the authored pivot table.</summary>
        public bool ShowRowGrandTotals { get; set; } = true;

        /// <summary>Gets or sets showColumnGrandTotals for the authored pivot table.</summary>
        public bool ShowColumnGrandTotals { get; set; } = true;

        /// <summary>Gets or sets pivotStyleName for the authored pivot table.</summary>
        public string? PivotStyleName { get; set; }

        /// <summary>Gets or sets layout for the authored pivot table.</summary>
        public ExcelPivotLayout Layout { get; set; } = ExcelPivotLayout.Compact;

        /// <summary>Gets or sets dataOnRows for the authored pivot table.</summary>
        public bool? DataOnRows { get; set; }

        /// <summary>Gets or sets showHeaders for the authored pivot table.</summary>
        public bool? ShowHeaders { get; set; }

        /// <summary>Gets or sets showEmptyRows for the authored pivot table.</summary>
        public bool? ShowEmptyRows { get; set; }

        /// <summary>Gets or sets showEmptyColumns for the authored pivot table.</summary>
        public bool? ShowEmptyColumns { get; set; }

        /// <summary>Gets or sets showDrill for the authored pivot table.</summary>
        public bool? ShowDrill { get; set; }

        /// <summary>Gets the fieldOptions used when creating the pivot table.</summary>
        public IList<ExcelPivotFieldOptions> FieldOptions { get; set; } = new List<ExcelPivotFieldOptions>();

        /// <summary>Gets or sets rowHeaderCaption for the authored pivot table.</summary>
        public string? RowHeaderCaption { get; set; }

        /// <summary>Gets or sets columnHeaderCaption for the authored pivot table.</summary>
        public string? ColumnHeaderCaption { get; set; }

        /// <summary>Gets or sets grandTotalCaption for the authored pivot table.</summary>
        public string? GrandTotalCaption { get; set; }

        /// <summary>Gets or sets missingCaption for the authored pivot table.</summary>
        public string? MissingCaption { get; set; }

        /// <summary>Gets or sets errorCaption for the authored pivot table.</summary>
        public string? ErrorCaption { get; set; }

        /// <summary>Gets or sets showDataDropDown for the authored pivot table.</summary>
        public bool? ShowDataDropDown { get; set; }

        /// <summary>Gets or sets showDropZones for the authored pivot table.</summary>
        public bool? ShowDropZones { get; set; }

        /// <summary>Gets or sets showDataTips for the authored pivot table.</summary>
        public bool? ShowDataTips { get; set; }

        /// <summary>Gets or sets showMemberPropertyTips for the authored pivot table.</summary>
        public bool? ShowMemberPropertyTips { get; set; }

        /// <summary>Gets or sets fieldListSortAscending for the authored pivot table.</summary>
        public bool? FieldListSortAscending { get; set; }

        /// <summary>Gets or sets customListSort for the authored pivot table.</summary>
        public bool? CustomListSort { get; set; }

        /// <summary>Gets the pivotFilters used when creating the pivot table.</summary>
        public IList<ExcelPivotFilter> PivotFilters { get; set; } = new List<ExcelPivotFilter>();

        /// <summary>Gets the calculatedFields used when creating the pivot table.</summary>
        public IList<ExcelPivotCalculatedField> CalculatedFields { get; set; } = new List<ExcelPivotCalculatedField>();

        /// <summary>Gets the groupings used when creating the pivot table.</summary>
        public IList<ExcelPivotGrouping> Groupings { get; set; } = new List<ExcelPivotGrouping>();
    }
}
