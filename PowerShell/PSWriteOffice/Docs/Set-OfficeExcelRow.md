---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Set-OfficeExcelRow
## SYNOPSIS
Writes values and applies row layout to a worksheet in a DSL block or through the pipeline.

## SYNTAX
### Default (Default)
```powershell
Set-OfficeExcelRow [-Row] <int> [[-Values] <Object[]>] [-Document <ExcelDocument>] [-Sheet <OfficeSheetSelector>] [-StartColumn <Object>] [-Height <ExcelColumnWidth>] [-ClearHeight] [-AutoFit] [-Hidden] [-Bold] [-Italic] [-Underline] [-WrapText] [-Font <string>] [-BackgroundColor <OfficeColor>] [-FirstColumn <Object>] [-LastColumn <Object>] [-Formula] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Writes values and applies row layout to a worksheet in a DSL block or through the pipeline.

## EXAMPLES

### EXAMPLE 1
```powershell
New-OfficeExcel -Path .\Rota.xlsx {
    ExcelSheet 'On call' {
        ExcelRow 1 'Week', 'Primary', 'Secondary', 'Escalation' -Bold
    }
}
```

Writes column labels across row 1 and makes them bold.

### EXAMPLE 2
```powershell
New-OfficeExcel -Path .\Reports\Budget-2027.xlsx {
    ExcelSheet 'Departments' {
        ExcelCell A1 'Budget 2027 by department'
        ExcelRow 3 'Department', 'Q1', 'Q2', 'Q3', 'Q4', 'Total' -Bold -BackgroundColor '#F1F5F9' -Height 22
    }
}
```

Writes the quarter labels into row 3 below the sheet title, with a light fill and a taller row so the band stands apart from the data.

### EXAMPLE 3
```powershell
Open-OfficeExcel -Path .\Finance\Budget-2027.xlsx {
    ExcelSheet 'Departments' {
        ExcelRow 20 'Total', '=SUM(B4:B19)', '=SUM(C4:C19)', '=SUM(D4:D19)', '=SUM(E4:E19)' -Formula -Bold
    }
}
```

Adds a totals row whose cells are formulas; without -Formula the same strings would be stored as text.

## PARAMETERS

### -AutoFit
Auto-fit row height after writing values and applying style.

```yaml
Type: SwitchParameter
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -BackgroundColor
Apply a background color across the target row span.

```yaml
Type: OfficeColor
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Bold
Apply or clear bold styling across the target row span.

```yaml
Type: SwitchParameter
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ClearHeight
Clear any custom row height.

```yaml
Type: SwitchParameter
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Document
Workbook to modify. Accepts pipeline input; select its worksheet with Sheet.

```yaml
Type: ExcelDocument
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -FirstColumn
First 1-based column affected by style options.

```yaml
Type: Object
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Font
Apply a font family across the target row span.

```yaml
Type: String
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Formula
Interpret explicitly prefixed row strings as formulas.

```yaml
Type: SwitchParameter
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Height
Explicit row height in points.

```yaml
Type: ExcelColumnWidth
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Hidden
Hide or unhide the row. Use -Hidden:$false to unhide.

```yaml
Type: SwitchParameter
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Italic
Apply or clear italic styling across the target row span.

```yaml
Type: SwitchParameter
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -LastColumn
Last 1-based column affected by style options.

```yaml
Type: Object
Parameter Sets: Default
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
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Row
1-based row index.

```yaml
Type: Int32
Parameter Sets: Default
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Sheet
Worksheet name when using Document. The first worksheet is used when no selector is supplied.

```yaml
Type: OfficeSheetSelector
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -StartColumn
Starting column index (1-based).

```yaml
Type: Object
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Underline
Apply or clear underline styling across the target row span.

```yaml
Type: SwitchParameter
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Values
Values to write across the row. Numeric, Boolean, and date values retain their native Excel types; strings remain text.

```yaml
Type: Object[]
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: 1
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -WrapText
Apply or clear wrap text across the target row span.

```yaml
Type: SwitchParameter
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
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
