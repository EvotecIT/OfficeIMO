---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Unprotect-OfficeExcelSheet
## SYNOPSIS
Removes worksheet protection.

## SYNTAX
### Default (Default)
```powershell
Unprotect-OfficeExcelSheet [-Document <ExcelDocument>] [-Sheet <OfficeSheetSelector>] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Removes worksheet protection.

## EXAMPLES

### EXAMPLE 1
```powershell
Open-OfficeExcel -Path .\Finance\Pricing-Model.xlsx {
    Unprotect-OfficeExcelSheet -Sheet Inputs
}
```

Unprotects the Inputs sheet.

### EXAMPLE 2
```powershell
Open-OfficeExcel -Path .\Finance\Pricing-Model.xlsx {
    ExcelSheet 'Rates' {
        Unprotect-OfficeExcelSheet
        ExcelCell B2 1.09
        ExcelCell B3 0.86
        Protect-OfficeExcelSheet -AllowSort
    }
}
```

Unprotects the price sheet, writes the new rates and restores protection with sorting allowed.

## PARAMETERS

### -Document
Caller-owned workbook; omission uses the enclosing Excel content block.

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

### -PassThru
Emits the native result after the operation.

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

### -Sheet
Worksheet name, one-based position or native worksheet object. Multi-sheet commands also accept wildcard names.

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

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `OfficeIMO.Excel.ExcelDocument`

## OUTPUTS

- `None`

## RELATED LINKS

- None
