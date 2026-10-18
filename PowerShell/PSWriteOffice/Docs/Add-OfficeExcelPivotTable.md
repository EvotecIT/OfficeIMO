---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Add-OfficeExcelPivotTable
## SYNOPSIS
Adds a pivot table to a worksheet.

## SYNTAX
### Context (Default)
```powershell
Add-OfficeExcelPivotTable -SourceRange <string> -DestinationCell <string> [-Name <string>] [-RowField <string[]>] [-ColumnField <string[]>] [-PageField <string[]>] [-DataField <string[]>] [-DataFunction <ExcelPivotDataFunction[]>] [-DataDisplayName <string[]>] [-DataNumberFormat <string[]>] [-NoRowGrandTotals] [-NoColumnGrandTotals] [-PivotStyle <string>] [-Layout <ExcelPivotLayout>] [-DataOnRows] [-DataOnColumns] [-ShowHeaders] [-HideHeaders] [-ShowEmptyRows] [-HideEmptyRows] [-ShowEmptyColumns] [-HideEmptyColumns] [-ShowDrill] [-HideDrill] [-RowHeaderCaption <string>] [-ColumnHeaderCaption <string>] [-GrandTotalCaption <string>] [-MissingCaption <string>] [-ErrorCaption <string>] [-ShowDataDropDown] [-HideDataDropDown] [-ShowDropZones] [-HideDropZones] [-ShowDataTips] [-HideDataTips] [-ShowMemberPropertyTips] [-HideMemberPropertyTips] [-FieldListSortAscending] [-FieldListSortDescending] [-CustomListSort] [-NoCustomListSort] [-FieldSort <hashtable>] [-FieldHiddenItems <hashtable>] [-FieldVisibleItems <hashtable>] [-PageFieldSelection <hashtable>] [-FieldNoDefaultSubtotal <string[]>] [-FieldSubtotalTop <string[]>] [-FieldInsertBlankRow <string[]>] [-FieldInsertPageBreak <string[]>] [-FieldCompact <string[]>] [-FieldOutline <string[]>] [-FieldHideDropDowns <string[]>] [-RefreshOnOpen] [-NoRefreshOnOpen] [-SaveSourceData] [-NoSaveSourceData] [-PreserveFormatting] [-NoPreserveFormatting] [-EnableDrill] [-DisableDrill] [-PassThru] [<CommonParameters>]
```

### Document
```powershell
Add-OfficeExcelPivotTable -Document <ExcelDocument> -SourceRange <string> -DestinationCell <string> [-Sheet <string>] [-SheetIndex <Int32>] [-Name <string>] [-RowField <string[]>] [-ColumnField <string[]>] [-PageField <string[]>] [-DataField <string[]>] [-DataFunction <ExcelPivotDataFunction[]>] [-DataDisplayName <string[]>] [-DataNumberFormat <string[]>] [-NoRowGrandTotals] [-NoColumnGrandTotals] [-PivotStyle <string>] [-Layout <ExcelPivotLayout>] [-DataOnRows] [-DataOnColumns] [-ShowHeaders] [-HideHeaders] [-ShowEmptyRows] [-HideEmptyRows] [-ShowEmptyColumns] [-HideEmptyColumns] [-ShowDrill] [-HideDrill] [-RowHeaderCaption <string>] [-ColumnHeaderCaption <string>] [-GrandTotalCaption <string>] [-MissingCaption <string>] [-ErrorCaption <string>] [-ShowDataDropDown] [-HideDataDropDown] [-ShowDropZones] [-HideDropZones] [-ShowDataTips] [-HideDataTips] [-ShowMemberPropertyTips] [-HideMemberPropertyTips] [-FieldListSortAscending] [-FieldListSortDescending] [-CustomListSort] [-NoCustomListSort] [-FieldSort <hashtable>] [-FieldHiddenItems <hashtable>] [-FieldVisibleItems <hashtable>] [-PageFieldSelection <hashtable>] [-FieldNoDefaultSubtotal <string[]>] [-FieldSubtotalTop <string[]>] [-FieldInsertBlankRow <string[]>] [-FieldInsertPageBreak <string[]>] [-FieldCompact <string[]>] [-FieldOutline <string[]>] [-FieldHideDropDowns <string[]>] [-RefreshOnOpen] [-NoRefreshOnOpen] [-SaveSourceData] [-NoSaveSourceData] [-PreserveFormatting] [-NoPreserveFormatting] [-EnableDrill] [-DisableDrill] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Adds a pivot table to a worksheet.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $rows = @(
    [pscustomobject]@{ Region = 'North America'; Product = 'Standard'; Sales = 125000 }
    [pscustomobject]@{ Region = 'EMEA'; Product = 'Standard'; Sales = 98000 }
    [pscustomobject]@{ Region = 'APAC'; Product = 'Premium'; Sales = 143000 }
)
New-OfficeExcel -Path .\SalesPivot.xlsx {
    Add-OfficeExcelSheet -Name Data {
        Add-OfficeExcelTable -InputObject $rows -TableName Sales -AutoFit
        Add-OfficeExcelPivotTable -SourceRange 'A1:C4' -DestinationCell 'E2' -Name 'SalesByRegion' -RowField Region -ColumnField Product -DataField Sales -DataFunction Sum -PivotStyle PivotStyleMedium9
    }
}
```

Writes source rows to a worksheet and creates a pivot table using the existing OfficeIMO pivot support.

## PARAMETERS

### -ColumnField
Column fields (header names).

```yaml
Type: String[]
Parameter Sets: Context, Document
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
Parameter Sets: Context, Document
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
Parameter Sets: Context, Document
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
Parameter Sets: Context, Document
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
Parameter Sets: Context, Document
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
Parameter Sets: Context, Document
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
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -DataOnColumns
Show data fields on columns.

```yaml
Type: SwitchParameter
Parameter Sets: Context, Document
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
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -DestinationCell
Top-left destination cell for the pivot table (e.g., "F2").

```yaml
Type: String
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -DisableDrill
Disable pivot detail drill interaction in Excel.

```yaml
Type: SwitchParameter
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Document
Workbook to operate on outside the DSL context.

```yaml
Type: ExcelDocument
Parameter Sets: Document
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -EnableDrill
Allow users to drill into pivot details in Excel.

```yaml
Type: SwitchParameter
Parameter Sets: Context, Document
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
Parameter Sets: Context, Document
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
Parameter Sets: Context, Document
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
Parameter Sets: Context, Document
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
Parameter Sets: Context, Document
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
Parameter Sets: Context, Document
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
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -FieldListSortAscending
Sort pivot field list ascending.

```yaml
Type: SwitchParameter
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -FieldListSortDescending
Sort pivot field list descending.

```yaml
Type: SwitchParameter
Parameter Sets: Context, Document
Aliases: None
Possible values:

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
Parameter Sets: Context, Document
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
Parameter Sets: Context, Document
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
Parameter Sets: Context, Document
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
Parameter Sets: Context, Document
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
Parameter Sets: Context, Document
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
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -HideDataDropDown
Hide the data drop-down.

```yaml
Type: SwitchParameter
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -HideDataTips
Hide pivot data tips.

```yaml
Type: SwitchParameter
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -HideDrill
Hide drill indicators.

```yaml
Type: SwitchParameter
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -HideDropZones
Hide pivot drop zones.

```yaml
Type: SwitchParameter
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -HideEmptyColumns
Hide empty columns.

```yaml
Type: SwitchParameter
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -HideEmptyRows
Hide empty rows.

```yaml
Type: SwitchParameter
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -HideHeaders
Hide field headers.

```yaml
Type: SwitchParameter
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -HideMemberPropertyTips
Hide member property tips.

```yaml
Type: SwitchParameter
Parameter Sets: Context, Document
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
Parameter Sets: Context, Document
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
Parameter Sets: Context, Document
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
Parameter Sets: Context, Document
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
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NoCustomListSort
Disable Excel custom-list sorting.

```yaml
Type: SwitchParameter
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NoPreserveFormatting
Do not preserve pivot formatting when Excel refreshes the pivot table.

```yaml
Type: SwitchParameter
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NoRefreshOnOpen
Do not refresh the pivot cache when the workbook opens.

```yaml
Type: SwitchParameter
Parameter Sets: Context, Document
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
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NoSaveSourceData
Do not save pivot source cache records in the workbook package.

```yaml
Type: SwitchParameter
Parameter Sets: Context, Document
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
Parameter Sets: Context, Document
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
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PassThru
Emit the worksheet after creating the pivot table.

```yaml
Type: SwitchParameter
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PivotStyle
Optional pivot table style name.

```yaml
Type: String
Parameter Sets: Context, Document
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
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -RefreshOnOpen
Refresh the pivot cache when the workbook opens.

```yaml
Type: SwitchParameter
Parameter Sets: Context, Document
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
Parameter Sets: Context, Document
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
Parameter Sets: Context, Document
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
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Sheet
Worksheet name when using Document.

```yaml
Type: String
Parameter Sets: Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -SheetIndex
Worksheet index (0-based) when using Document.

```yaml
Type: Int32
Parameter Sets: Document
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
Parameter Sets: Context, Document
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
Parameter Sets: Context, Document
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
Parameter Sets: Context, Document
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
Parameter Sets: Context, Document
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
Parameter Sets: Context, Document
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
Parameter Sets: Context, Document
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
Parameter Sets: Context, Document
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
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -SourceRange
Source data range including header row (e.g., "A1:D200").

```yaml
Type: String
Parameter Sets: Context, Document
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
