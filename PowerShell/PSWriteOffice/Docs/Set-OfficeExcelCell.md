---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Set-OfficeExcelCell
## SYNOPSIS
Sets a cell value, formula, rich text, number format or fill.

## SYNTAX
### Address (Default)
```powershell
Set-OfficeExcelCell [-Address] <string> [[-Value] <Object>] [-Formula <string>] [-Document <ExcelDocument>] [-Sheet <OfficeSheetSelector>] [-NumberFormat <string>] [-BackgroundColor <OfficeColor>] [-GradientFrom <OfficeColor>] [-GradientTo <OfficeColor>] [-GradientAngle <Double>] [-PassThru] [<CommonParameters>]
```

### Coordinates
```powershell
Set-OfficeExcelCell [[-Value] <Object>] -Row <Int32> -Column <Object> [-Formula <string>] [-Document <ExcelDocument>] [-Sheet <OfficeSheetSelector>] [-NumberFormat <string>] [-BackgroundColor <OfficeColor>] [-GradientFrom <OfficeColor>] [-GradientTo <OfficeColor>] [-GradientAngle <Double>] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Sets a cell value, formula, rich text, number format or fill.

## EXAMPLES

### EXAMPLE 1
```powershell
New-OfficeExcel -Path .\Budget.xlsx {
    ExcelSheet 'Summary' {
        ExcelCell A1 'Budget 2027'
    }
}
```

Writes a heading into A1 of the Summary sheet.

### EXAMPLE 2
```powershell
New-OfficeExcel -Path .\Estimates\Migration-Cost.xlsx {
    ExcelSheet 'Estimate' {
        ExcelCell A1 'Servers'
        ExcelCell B1 42
        ExcelCell A2 'Cost per server'
        ExcelCell B2 1850 -NumberFormat '#,##0.00'
        ExcelCell A3 'Total'
        ExcelCell B3 -Formula '=B1*B2' -NumberFormat '#,##0.00' -BackgroundColor '#FEF3C7'
    }
}
```

Writes labels, inputs and formulas for a cost estimate, formatting the result as currency.

### EXAMPLE 3
```powershell
Open-OfficeExcel -Path .\Reports\Status.xlsx {
    ExcelSheet 'Summary' {
        ExcelCell B2 {
            ExcelText 'Status: ' -Bold
            ExcelText 'On track' -Bold -Color Success
        }
    }
}
```

Passes a block of text runs as the value, so one cell mixes bold and coloured text.

## PARAMETERS

### -Address
A1 address of the target cell.

```yaml
Type: String
Parameter Sets: Address
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -BackgroundColor
Cell fill colour expressed as a literal OfficeColor or theme role.

```yaml
Type: OfficeColor
Parameter Sets: Address, Coordinates
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
Type: Object
Parameter Sets: Coordinates
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Document
Caller-owned workbook; omission uses the enclosing Excel content block.

```yaml
Type: ExcelDocument
Parameter Sets: Address, Coordinates
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -Formula
Formula text (without leading =).

```yaml
Type: String
Parameter Sets: Address, Coordinates
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -GradientAngle
Linear fill angle in degrees.

```yaml
Type: Double
Parameter Sets: Address, Coordinates
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -GradientFrom
First fill colour; GradientTo is required with it.

```yaml
Type: OfficeColor
Parameter Sets: Address, Coordinates
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -GradientTo
Second fill colour; GradientFrom is required with it.

```yaml
Type: OfficeColor
Parameter Sets: Address, Coordinates
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NumberFormat
Number format code to apply.

```yaml
Type: String
Parameter Sets: Address, Coordinates
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
Parameter Sets: Address, Coordinates
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
Parameter Sets: Coordinates
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Sheet
Worksheet name, one-based position or native worksheet object. Multi-sheet commands also accept wildcard names.

```yaml
Type: OfficeSheetSelector
Parameter Sets: Address, Coordinates
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Value
Value or values consumed by the selected operation.

```yaml
Type: Object
Parameter Sets: Address, Coordinates
Aliases: None
Possible values:

Required: False
Position: 1
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
