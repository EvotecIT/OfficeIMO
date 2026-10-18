---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Add-OfficeExcelSubtotalSummary
## SYNOPSIS
Adds grouped subtotal summary rows for a worksheet data range.

## SYNTAX
### Context (Default)
```powershell
Add-OfficeExcelSubtotalSummary -GroupColumn <string> -ValueColumn <string[]> [-HeaderRow <Int32>] [-DataStartRow <Int32>] [-DataEndRow <Int32>] [-SummaryStartRow <Int32>] [-Function <ExcelSubtotalFunction>] [-NoHeader] [-NoGrandTotal] [-NoOutline] [-HideDetailRows] [-OutlineLevel <int>] [-LabelSuffix <string>] [-GrandTotalLabel <string>] [-PassThru] [<CommonParameters>]
```

### Document
```powershell
Add-OfficeExcelSubtotalSummary -Document <ExcelDocument> -GroupColumn <string> -ValueColumn <string[]> [-Sheet <string>] [-SheetIndex <Int32>] [-HeaderRow <Int32>] [-DataStartRow <Int32>] [-DataEndRow <Int32>] [-SummaryStartRow <Int32>] [-Function <ExcelSubtotalFunction>] [-NoHeader] [-NoGrandTotal] [-NoOutline] [-HideDetailRows] [-OutlineLevel <int>] [-LabelSuffix <string>] [-GrandTotalLabel <string>] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Adds grouped subtotal summary rows for a worksheet data range.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> ExcelSheet Data { Add-OfficeExcelSubtotalSummary -GroupColumn Region -ValueColumn Sales -DataEndRow 20 }
```

Writes SUBTOTAL formulas below the data range and applies row outline metadata to each group.

## PARAMETERS

### -DataEndRow
Last data row. Defaults to the last row of the used range.

```yaml
Type: Int32
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -DataStartRow
First data row. Defaults to the row after HeaderRow.

```yaml
Type: Int32
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

### -Function
Subtotal function.

```yaml
Type: ExcelSubtotalFunction
Parameter Sets: Context, Document
Aliases: None
Possible values: Average, Count, CountNonBlank, Max, Min, Sum

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -GrandTotalLabel
Label used for the optional grand total row.

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

### -GroupColumn
Group column as a 1-based index, column letter, or header name.

```yaml
Type: String
Parameter Sets: Context, Document
Aliases: By, GroupBy
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -HeaderRow
Header row that contains source labels. Defaults to the first row of the used range.

```yaml
Type: Int32
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -HideDetailRows
Hide detail rows when applying outline metadata.

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

### -LabelSuffix
Text appended to each group key in the subtotal label cell.

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

### -NoGrandTotal
Skip writing a grand total row.

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

### -NoHeader
Skip writing a summary header row.

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

### -NoOutline
Skip applying outline metadata to detail rows.

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

### -OutlineLevel
Outline level used for grouped detail rows.

```yaml
Type: Int32
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
Emit OfficeIMO subtotal generation metadata.

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

### -SummaryStartRow
First row for the generated summary block.

```yaml
Type: Int32
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ValueColumn
Value columns as 1-based indexes, column letters, or header names.

```yaml
Type: String[]
Parameter Sets: Context, Document
Aliases: ValueColumns, AggregateColumn, AggregateColumns
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

- `OfficeIMO.Excel.ExcelSubtotalResult`

## RELATED LINKS

- None
