---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Set-OfficeExcelColumn
## SYNOPSIS
Writes values or applies column layout to a worksheet in a DSL block or through the pipeline.

## SYNTAX
### __AllParameterSets
```powershell
Set-OfficeExcelColumn [[-Column] <Int32>] [-Worksheet <ExcelSheet>] [-Document <ExcelDocument>] [-Sheet <string>] [-SheetIndex <Int32>] [-ColumnName <string>] [-Values <Object[]>] [-StartRow <int>] [-Width <Double>] [-Hidden <Boolean>] [-AutoFit] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Writes values or applies column layout to a worksheet in a DSL block or through the pipeline.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> ExcelSheet 'Data' { Set-OfficeExcelColumn -Column 1 -Values 'North','South' -AutoFit }
```

Writes values into column A and adjusts the width.

### EXAMPLE 2
```powershell
PS> $sheet | Set-OfficeExcelColumn -ColumnName C -Width 48
```

Sets column C to 48 Excel character units. Save the workbook after editing.

### EXAMPLE 3
```powershell
PS> $workbook | Set-OfficeExcelColumn -Sheet 'Review' -ColumnName C -Hidden $false -AutoFit
```

Uses the workbook pipeline target without an active DSL scope.

## PARAMETERS

### -AutoFit
Auto-fit the column width after updates.

```yaml
Type: SwitchParameter
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Column
1-based column index.

```yaml
Type: Int32
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ColumnName
Column letter reference (e.g., A, BC).

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: ColumnLetter, Letter
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Document
Workbook to modify. Accepts pipeline input; select its worksheet with Sheet or SheetIndex.

```yaml
Type: ExcelDocument
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -Hidden
Hide or show the column.

```yaml
Type: Boolean
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PassThru
Emit the object created or changed by the command.

```yaml
Type: SwitchParameter
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Sheet
Worksheet name when using Document. The first worksheet is used when no selector is supplied.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -SheetIndex
Zero-based worksheet index when using Document.

```yaml
Type: Int32
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -StartRow
Starting row index (1-based) for values.

```yaml
Type: Int32
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Values
Values to write down the column. Numeric, Boolean, and date values retain their native Excel types; strings remain text.

```yaml
Type: Object[]
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Width
Column width in Excel character units.

```yaml
Type: Double
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Worksheet
Worksheet to modify outside a DSL context.

```yaml
Type: ExcelSheet
Parameter Sets: __AllParameterSets
Aliases: SheetObject
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `OfficeIMO.Excel.ExcelSheet`
- `OfficeIMO.Excel.ExcelDocument`

## OUTPUTS

- `None`

## RELATED LINKS

- None
