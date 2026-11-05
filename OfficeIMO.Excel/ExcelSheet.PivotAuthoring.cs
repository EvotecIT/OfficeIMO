namespace OfficeIMO.Excel;

public partial class ExcelSheet {
    /// <summary>Creates a pivot table from a worksheet range and returns its editable object.</summary>
    public ExcelPivotTable AddPivotTable(string sourceRange, string destinationCell) => AddPivotTable(sourceRange, destinationCell, null);

    /// <summary>Creates a pivot table from a worksheet range using the supplied authoring options.</summary>
    public ExcelPivotTable AddPivotTable(string sourceRange, string destinationCell, ExcelPivotTableOptions? options) {
        return AddPivotTable(this, sourceRange, destinationCell, options);
    }

    /// <summary>Creates an editable pivot on this worksheet using source cells from another worksheet in the same workbook.</summary>
    public ExcelPivotTable AddPivotTable(ExcelSheet sourceSheet, string sourceRange, string destinationCell, ExcelPivotTableOptions? options = null) {
        ArgumentNullException.ThrowIfNull(sourceSheet);
        if (!ReferenceEquals(_excelDocument, sourceSheet._excelDocument))
            throw new ArgumentException("The pivot source worksheet must belong to the same workbook.", nameof(sourceSheet));
        var worksheets = WorkbookPartRoot.Workbook?.Sheets;
        if (worksheets == null || !ReferenceEquals(sourceSheet.SheetElement.Parent, worksheets))
            throw new ArgumentException("The pivot source worksheet is no longer attached to this workbook.", nameof(sourceSheet));
        options ??= new ExcelPivotTableOptions();
        return AddPivotTableCore(
            sourceSheet: sourceSheet,
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
