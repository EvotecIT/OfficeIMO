---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Set-OfficeExcelDateSystem
## SYNOPSIS
Sets the workbook date system (1900 or 1904).

## SYNTAX
### Default (Default)
```powershell
Set-OfficeExcelDateSystem [-DateSystem] <ExcelDateSystem> [-Document <ExcelDocument>] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Sets the workbook date system (1900 or 1904).

## EXAMPLES

### EXAMPLE 1
```powershell
Open-OfficeExcel -Path .\Finance\Legacy-Ledger.xlsx {
    ExcelDateSystem 1904
}
```

Switches the workbook to the 1904 date system used by some older Mac workbooks.

### EXAMPLE 2
```powershell
New-OfficeExcel -Path .\Finance\Consolidated.xlsx {
    ExcelDateSystem 1904
    Copy-OfficeExcelSheet -Source .\Partners\Ledger-1904.xlsx -Sheet Ledger
}
```

Sets the date system of a new workbook to match a 1904-based partner file before copying its sheets in.

## PARAMETERS

### -DateSystem
Date system to use for Excel date serials.

```yaml
Type: ExcelDateSystem
Parameter Sets: Default
Aliases: None
Possible values: NineteenHundred, NineteenFour

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

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

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `OfficeIMO.Excel.ExcelDocument`

## OUTPUTS

- `None`

## RELATED LINKS

- None
