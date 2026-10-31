namespace OfficeIMO.Excel;

public partial class ExcelSheet {
    /// <summary>Creates a pivot table from a worksheet range and returns its editable object.</summary>
    public ExcelPivotTable AddPivotTable(string sourceRange, string destinationCell, ExcelPivotTableOptions? options = null) {
        options ??= new ExcelPivotTableOptions();
        return AddPivotTableCore(
            sourceRange: sourceRange,
            destinationCell: destinationCell,
            name: options.Name,
            rowFields: options.RowFields,
            columnFields: options.ColumnFields,
            pageFields: options.PageFields,
            dataFields: options.DataFields,
            showRowGrandTotals: options.ShowRowGrandTotals,
            showColumnGrandTotals: options.ShowColumnGrandTotals,
            pivotStyleName: options.PivotStyleName,
            layout: options.Layout,
            dataOnRows: options.DataOnRows,
            showHeaders: options.ShowHeaders,
            showEmptyRows: options.ShowEmptyRows,
            showEmptyColumns: options.ShowEmptyColumns,
            showDrill: options.ShowDrill,
            fieldOptions: options.FieldOptions,
            rowHeaderCaption: options.RowHeaderCaption,
            columnHeaderCaption: options.ColumnHeaderCaption,
            grandTotalCaption: options.GrandTotalCaption,
            missingCaption: options.MissingCaption,
            errorCaption: options.ErrorCaption,
            showDataDropDown: options.ShowDataDropDown,
            showDropZones: options.ShowDropZones,
            showDataTips: options.ShowDataTips,
            showMemberPropertyTips: options.ShowMemberPropertyTips,
            fieldListSortAscending: options.FieldListSortAscending,
            customListSort: options.CustomListSort,
            pivotFilters: options.PivotFilters,
            calculatedFields: options.CalculatedFields,
            groupings: options.Groupings,
            options: options);
    }
}
