---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Rename-OfficeExcelNamedRange
## SYNOPSIS
Renames a defined name.

## SYNTAX
### Default (Default)
```powershell
Rename-OfficeExcelNamedRange [-Name] <string> [-NewName] <string> [-Document <ExcelDocument>] [-Sheet <OfficeSheetSelector>] [-Global] [-ValidationMode <ExcelDefinedNameValidationMode>] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Renames a defined name.

## EXAMPLES

### EXAMPLE 1
```powershell
Open-OfficeExcel -Path .\Finance\Budget-2027.xlsx {
    Rename-OfficeExcelNamedRange Rate TaxRate
}
```

Renames Rate to TaxRate; formulas that use the name follow the change.

### EXAMPLE 2
```powershell
$names = Get-OfficeExcelNamedRange -Path .\Finance\Pricing-Model.xlsx | Where-Object SheetName -eq 'Inputs'
Open-OfficeExcel -Path .\Finance\Pricing-Model.xlsx {
    foreach ($name in $names) {
        Rename-OfficeExcelNamedRange -Name $name.Name -NewName "in_$($name.Name)" -ValidationMode Strict
    }
}
```

Prefixes every input name of the pricing model with in_ so inputs are easy to tell apart from calculations.

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

### -Global
Use workbook-global scope from inside the DSL.

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

### -Name
Named range name.

```yaml
Type: String
Parameter Sets: Default
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NewName
New named range name.

```yaml
Type: String
Parameter Sets: Default
Aliases: None
Possible values:

Required: True
Position: 1
Default value: None
Accept pipeline input: False
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

### -ValidationMode
Defined-name validation mode.

```yaml
Type: ExcelDefinedNameValidationMode
Parameter Sets: Default
Aliases: None
Possible values: Sanitize, Strict

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
