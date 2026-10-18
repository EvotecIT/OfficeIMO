---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Set-OfficeExcelPageSetup
## SYNOPSIS
Configures page setup options on a worksheet.

## SYNTAX
### Context (Default)
```powershell
Set-OfficeExcelPageSetup [-FitToWidth <UInt32>] [-FitToHeight <UInt32>] [-Scale <UInt32>] [-PageOrder <ExcelPageOrder>] [-PassThru] [<CommonParameters>]
```

### Document
```powershell
Set-OfficeExcelPageSetup -Document <ExcelDocument> [-Sheet <string>] [-SheetIndex <Int32>] [-FitToWidth <UInt32>] [-FitToHeight <UInt32>] [-Scale <UInt32>] [-PageOrder <ExcelPageOrder>] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Configures page setup options on a worksheet.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> ExcelSheet 'Data' { Set-OfficeExcelPageSetup -FitToWidth 1 -FitToHeight 0 }
```

Fits the sheet to one page wide and unlimited height.

## PARAMETERS

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

### -FitToHeight
Number of pages to fit vertically.

```yaml
Type: UInt32
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -FitToWidth
Number of pages to fit horizontally.

```yaml
Type: UInt32
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PageOrder
Multi-page print order.

```yaml
Type: ExcelPageOrder
Parameter Sets: Context, Document
Aliases: None
Possible values: DownThenOver, OverThenDown

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PassThru
Emit the worksheet after applying settings.

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

### -Scale
Manual scale percentage (10-400).

```yaml
Type: UInt32
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

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `OfficeIMO.Excel.ExcelDocument`

## OUTPUTS

- `None`

## RELATED LINKS

- None
