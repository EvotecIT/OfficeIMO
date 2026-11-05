---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Add-OfficeExcelPivotTable
## SYNOPSIS
Adds a pivot table to a worksheet.

## SYNTAX
### Range (Default)
```powershell
Add-OfficeExcelPivotTable -Range <string> -Address <string> [-Document <ExcelDocument>] [-Sheet <OfficeSheetSelector>] [-SourceSheet <OfficeSheetSelector>] [-Name <string>] [-RowField <string[]>] [-ColumnField <string[]>] [-PageField <string[]>] [-DataField <string[]>] [-DataFunction <ExcelPivotDataFunction[]>] [-DataDisplayName <string[]>] [-DataNumberFormat <string[]>] [-NoRowGrandTotals] [-NoColumnGrandTotals] [-Style <string>] [-Layout <ExcelPivotLayout>] [-DataOnRows] [-ShowHeaders] [-ShowEmptyRows] [-ShowEmptyColumns] [-ShowDrill] [-EnableDrill] [-RowHeaderCaption <string>] [-ColumnHeaderCaption <string>] [-GrandTotalCaption <string>] [-MissingCaption <string>] [-ErrorCaption <string>] [-ShowDataDropDown] [-ShowDropZones] [-ShowDataTips] [-ShowMemberPropertyTips] [-FieldListSort <ListSortDirection>] [-CustomListSort] [-FieldSort <hashtable>] [-FieldHiddenItems <hashtable>] [-FieldVisibleItems <hashtable>] [-PageFieldSelection <hashtable>] [-FieldNoDefaultSubtotal <string[]>] [-FieldSubtotalTop <string[]>] [-FieldInsertBlankRow <string[]>] [-FieldInsertPageBreak <string[]>] [-FieldCompact <string[]>] [-FieldOutline <string[]>] [-FieldHideDropDowns <string[]>] [-RefreshOnOpen] [-SaveSourceData] [-PreserveFormatting] [-PassThru] [<CommonParameters>]
```

### SourceSheet
```powershell
Add-OfficeExcelPivotTable -SourceSheet <OfficeSheetSelector> -Address <string> [-Document <ExcelDocument>] [-Sheet <OfficeSheetSelector>] [-Name <string>] [-RowField <string[]>] [-ColumnField <string[]>] [-PageField <string[]>] [-DataField <string[]>] [-DataFunction <ExcelPivotDataFunction[]>] [-DataDisplayName <string[]>] [-DataNumberFormat <string[]>] [-NoRowGrandTotals] [-NoColumnGrandTotals] [-Style <string>] [-Layout <ExcelPivotLayout>] [-DataOnRows] [-ShowHeaders] [-ShowEmptyRows] [-ShowEmptyColumns] [-ShowDrill] [-EnableDrill] [-RowHeaderCaption <string>] [-ColumnHeaderCaption <string>] [-GrandTotalCaption <string>] [-MissingCaption <string>] [-ErrorCaption <string>] [-ShowDataDropDown] [-ShowDropZones] [-ShowDataTips] [-ShowMemberPropertyTips] [-FieldListSort <ListSortDirection>] [-CustomListSort] [-FieldSort <hashtable>] [-FieldHiddenItems <hashtable>] [-FieldVisibleItems <hashtable>] [-PageFieldSelection <hashtable>] [-FieldNoDefaultSubtotal <string[]>] [-FieldSubtotalTop <string[]>] [-FieldInsertBlankRow <string[]>] [-FieldInsertPageBreak <string[]>] [-FieldCompact <string[]>] [-FieldOutline <string[]>] [-FieldHideDropDowns <string[]>] [-RefreshOnOpen] [-SaveSourceData] [-PreserveFormatting] [-PassThru] [<CommonParameters>]
```

### Table
```powershell
Add-OfficeExcelPivotTable -Address <string> -Table <string> [-Document <ExcelDocument>] [-Sheet <OfficeSheetSelector>] [-Name <string>] [-RowField <string[]>] [-ColumnField <string[]>] [-PageField <string[]>] [-DataField <string[]>] [-DataFunction <ExcelPivotDataFunction[]>] [-DataDisplayName <string[]>] [-DataNumberFormat <string[]>] [-NoRowGrandTotals] [-NoColumnGrandTotals] [-Style <string>] [-Layout <ExcelPivotLayout>] [-DataOnRows] [-ShowHeaders] [-ShowEmptyRows] [-ShowEmptyColumns] [-ShowDrill] [-EnableDrill] [-RowHeaderCaption <string>] [-ColumnHeaderCaption <string>] [-GrandTotalCaption <string>] [-MissingCaption <string>] [-ErrorCaption <string>] [-ShowDataDropDown] [-ShowDropZones] [-ShowDataTips] [-ShowMemberPropertyTips] [-FieldListSort <ListSortDirection>] [-CustomListSort] [-FieldSort <hashtable>] [-FieldHiddenItems <hashtable>] [-FieldVisibleItems <hashtable>] [-PageFieldSelection <hashtable>] [-FieldNoDefaultSubtotal <string[]>] [-FieldSubtotalTop <string[]>] [-FieldInsertBlankRow <string[]>] [-FieldInsertPageBreak <string[]>] [-FieldCompact <string[]>] [-FieldOutline <string[]>] [-FieldHideDropDowns <string[]>] [-RefreshOnOpen] [-SaveSourceData] [-PreserveFormatting] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Adds a pivot table to a worksheet.

## EXAMPLES

### EXAMPLE 1
```powershell
Open-OfficeExcel -Path .\Reports\Q3-Sales.xlsx {
    ExcelSheet 'Orders' {
        ExcelPivotTable -Range A1:F500 -Address H2 -RowField Region -DataField Amount -DataFunction Sum
    }
}
```

Builds a pivot table of revenue by region from the order data in A1:F500.

### EXAMPLE 2
```powershell
$orders = Import-OfficeCsv -Path .\Data\Orders-2026.csv
New-OfficeExcel -Path .\Reports\Sales-2026.xlsx {
    ExcelSheet 'Orders' {
        ExcelTable $orders -Name Orders
        ExcelPivotTable -Table Orders -Address J3 -Name SalesPivot -RowField Region -ColumnField Quarter -PageField Product -DataField Amount -DataFunction Sum -DataDisplayName 'Revenue' -DataNumberFormat '#,##0' -Layout Tabular -RefreshOnOpen
    }
}
```

Writes a year of orders and pivots them by region and quarter with a product filter, naming the pivot so a slicer and conditional formats can refer to it.

### EXAMPLE 3
```powershell
Open-OfficeExcel -Path .\Reports\Q3-Sales.xlsx {
    ExcelSheet 'Summary' {
        ExcelPivotTable -SourceSheet Orders -Address A3 -RowField Region -ColumnField Product -DataField Amount -DataFunction Sum
    }
}
```

Builds a summary sheet whose pivot reads all the data on the Orders sheet.

## PARAMETERS

### -Address
A1 address of the target cell.

```yaml
Type: String
Parameter Sets: Range, SourceSheet, Table
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ColumnField
Column fields (header names).

```yaml
Type: String[]
Parameter Sets: Range, SourceSheet, Table
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ColumnHeaderCaption
Optional column header caption.

```yaml
Type: String
Parameter Sets: Range, SourceSheet, Table
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -CustomListSort
Use Excel custom-list sorting.

```yaml
Type: SwitchParameter
Parameter Sets: Range, SourceSheet, Table
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -DataDisplayName
Display names for data fields.

```yaml
Type: String[]
Parameter Sets: Range, SourceSheet, Table
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -DataField
Data fields (header names). Defaults to the last column when omitted.

```yaml
Type: String[]
Parameter Sets: Range, SourceSheet, Table
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -DataFunction
Aggregation functions (Sum, Count, Average, etc.).

```yaml
Type: ExcelPivotDataFunction[]
Parameter Sets: Range, SourceSheet, Table
Aliases: None
Possible values: Average, Count, CountNumbers, Maximum, Minimum, Product, StandardDeviation, StandardDeviationP, Sum, Variance, VarianceP

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -DataNumberFormat
Number format codes for data fields.

```yaml
Type: String[]
Parameter Sets: Range, SourceSheet, Table
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -DataOnRows
Show data fields on rows.

```yaml
Type: SwitchParameter
Parameter Sets: Range, SourceSheet, Table
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Document
Caller-owned workbook; omission uses the enclosing Excel content block.

```yaml
Type: ExcelDocument
Parameter Sets: Range, SourceSheet, Table
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -EnableDrill
Allow users to drill into pivot details in Excel.

```yaml
Type: SwitchParameter
Parameter Sets: Range, SourceSheet, Table
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ErrorCaption
Optional error-value caption.

```yaml
Type: String
Parameter Sets: Range, SourceSheet, Table
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -FieldCompact
Fields using compact field layout.

```yaml
Type: String[]
Parameter Sets: Range, SourceSheet, Table
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -FieldHiddenItems
Field item captions to hide, for example @{ Region = @('Legacy') }.

```yaml
Type: Hashtable
Parameter Sets: Range, SourceSheet, Table
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -FieldHideDropDowns
Fields whose filter drop-downs should be hidden.

```yaml
Type: String[]
Parameter Sets: Range, SourceSheet, Table
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -FieldInsertBlankRow
Fields that insert blank rows after items.

```yaml
Type: String[]
Parameter Sets: Range, SourceSheet, Table
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -FieldInsertPageBreak
Fields that insert page breaks after items.

```yaml
Type: String[]
Parameter Sets: Range, SourceSheet, Table
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -FieldListSort
Controls the ordering of the pivot field list.

```yaml
Type: ListSortDirection
Parameter Sets: Range, SourceSheet, Table
Aliases: None
Possible values: Ascending, Descending

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -FieldNoDefaultSubtotal
Fields with default subtotal disabled.

```yaml
Type: String[]
Parameter Sets: Range, SourceSheet, Table
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -FieldOutline
Fields using outline field layout.

```yaml
Type: String[]
Parameter Sets: Range, SourceSheet, Table
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -FieldSort
Field sort map, for example @{ Region = 'Ascending' }.

```yaml
Type: Hashtable
Parameter Sets: Range, SourceSheet, Table
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -FieldSubtotalTop
Fields with subtotals shown at the top.

```yaml
Type: String[]
Parameter Sets: Range, SourceSheet, Table
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -FieldVisibleItems
Field item captions to keep visible, hiding other known items.

```yaml
Type: Hashtable
Parameter Sets: Range, SourceSheet, Table
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -GrandTotalCaption
Optional grand total caption.

```yaml
Type: String
Parameter Sets: Range, SourceSheet, Table
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Layout
Pivot layout (Compact, Outline, Tabular).

```yaml
Type: ExcelPivotLayout
Parameter Sets: Range, SourceSheet, Table
Aliases: None
Possible values: Compact, Outline, Tabular

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -MissingCaption
Optional missing-value caption.

```yaml
Type: String
Parameter Sets: Range, SourceSheet, Table
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Name
Optional pivot table name.

```yaml
Type: String
Parameter Sets: Range, SourceSheet, Table
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NoColumnGrandTotals
Disable column grand totals.

```yaml
Type: SwitchParameter
Parameter Sets: Range, SourceSheet, Table
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NoRowGrandTotals
Disable row grand totals.

```yaml
Type: SwitchParameter
Parameter Sets: Range, SourceSheet, Table
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PageField
Page fields (header names) used as filters.

```yaml
Type: String[]
Parameter Sets: Range, SourceSheet, Table
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PageFieldSelection
Selected page-field item captions, for example @{ Product = 'Standard' }.

```yaml
Type: Hashtable
Parameter Sets: Range, SourceSheet, Table
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PassThru
Emits the native result after the operation.

```yaml
Type: SwitchParameter
Parameter Sets: Range, SourceSheet, Table
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PreserveFormatting
Preserve pivot formatting when Excel refreshes the pivot table.

```yaml
Type: SwitchParameter
Parameter Sets: Range, SourceSheet, Table
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Range
A1 cell range on the selected worksheet.

```yaml
Type: String
Parameter Sets: Range
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -RefreshOnOpen
Refresh the pivot cache when the workbook opens.

```yaml
Type: SwitchParameter
Parameter Sets: Range, SourceSheet, Table
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -RowField
Row fields (header names).

```yaml
Type: String[]
Parameter Sets: Range, SourceSheet, Table
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -RowHeaderCaption
Optional row header caption.

```yaml
Type: String
Parameter Sets: Range, SourceSheet, Table
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -SaveSourceData
Save pivot source cache records in the workbook package.

```yaml
Type: SwitchParameter
Parameter Sets: Range, SourceSheet, Table
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Sheet
Worksheet name, one-based position or native worksheet object. Multi-sheet commands also accept wildcard names.

```yaml
Type: OfficeSheetSelector
Parameter Sets: Range, SourceSheet, Table
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ShowDataDropDown
Show the data drop-down.

```yaml
Type: SwitchParameter
Parameter Sets: Range, SourceSheet, Table
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ShowDataTips
Show pivot data tips.

```yaml
Type: SwitchParameter
Parameter Sets: Range, SourceSheet, Table
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ShowDrill
Show drill indicators.

```yaml
Type: SwitchParameter
Parameter Sets: Range, SourceSheet, Table
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ShowDropZones
Show pivot drop zones.

```yaml
Type: SwitchParameter
Parameter Sets: Range, SourceSheet, Table
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ShowEmptyColumns
Show empty columns.

```yaml
Type: SwitchParameter
Parameter Sets: Range, SourceSheet, Table
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ShowEmptyRows
Show empty rows.

```yaml
Type: SwitchParameter
Parameter Sets: Range, SourceSheet, Table
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ShowHeaders
Show field headers.

```yaml
Type: SwitchParameter
Parameter Sets: Range, SourceSheet, Table
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ShowMemberPropertyTips
Show member property tips.

```yaml
Type: SwitchParameter
Parameter Sets: Range, SourceSheet, Table
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -SourceSheet
Gets or sets the source worksheet name, one-based position or native worksheet.

```yaml
Type: OfficeSheetSelector
Parameter Sets: Range, SourceSheet
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Style
Built-in chart style identifier.

```yaml
Type: String
Parameter Sets: Range, SourceSheet, Table
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Table
Native Excel table used to resolve the source range or header.

```yaml
Type: String
Parameter Sets: Table
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `OfficeIMO.Excel.ExcelDocument`

## OUTPUTS

- `None`

## RELATED LINKS

- None
