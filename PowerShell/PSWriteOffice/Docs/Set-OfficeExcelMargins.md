---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Set-OfficeExcelMargins
## SYNOPSIS
Sets page margins on a worksheet.

## SYNTAX
### Context (Default)
```powershell
Set-OfficeExcelMargins [-Preset <ExcelMarginPreset>] [-Left <Double>] [-Right <Double>] [-Top <Double>] [-Bottom <Double>] [-Header <Double>] [-Footer <Double>] [-PassThru] [<CommonParameters>]
```

### Document
```powershell
Set-OfficeExcelMargins -Document <ExcelDocument> [-Sheet <string>] [-SheetIndex <Int32>] [-Preset <ExcelMarginPreset>] [-Left <Double>] [-Right <Double>] [-Top <Double>] [-Bottom <Double>] [-Header <Double>] [-Footer <Double>] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Sets page margins on a worksheet.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> ExcelSheet 'Data' { Set-OfficeExcelMargins -Preset Narrow }
```

Applies the Narrow margin preset.

### EXAMPLE 2
```powershell
PS> ExcelSheet 'Data' { Set-OfficeExcelMargins -Left 0.5 -Right 0.5 -Top 0.75 -Bottom 0.75 }
```

Sets custom margins in inches.

## PARAMETERS

### -Bottom
Bottom margin in inches.

```yaml
Type: Double
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

### -Footer
Footer margin in inches.

```yaml
Type: Double
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Header
Header margin in inches.

```yaml
Type: Double
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Left
Left margin in inches.

```yaml
Type: Double
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
Emit the worksheet after applying margins.

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

### -Preset
Margin preset to apply.

```yaml
Type: ExcelMarginPreset
Parameter Sets: Context, Document
Aliases: None
Possible values: Normal, Narrow, Moderate, Wide

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Right
Right margin in inches.

```yaml
Type: Double
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

### -Top
Top margin in inches.

```yaml
Type: Double
Parameter Sets: Context, Document
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
