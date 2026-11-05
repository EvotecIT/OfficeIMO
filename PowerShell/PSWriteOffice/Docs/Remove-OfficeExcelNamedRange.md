---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Remove-OfficeExcelNamedRange
## SYNOPSIS
Removes a defined name.

## SYNTAX
### Default (Default)
```powershell
Remove-OfficeExcelNamedRange [-Name] <string> [-Document <ExcelDocument>] [-Sheet <OfficeSheetSelector>] [-Global] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Removes a defined name.

## EXAMPLES

### EXAMPLE 1
```powershell
Open-OfficeExcel -Path .\Finance\Budget-2027.xlsx {
    Remove-OfficeExcelNamedRange OldTaxRate
}
```

Deletes the OldTaxRate name from the workbook.

### EXAMPLE 2
```powershell
$broken = Get-OfficeExcelNamedRange -Path .\Finance\Budget-2027.xlsx | Where-Object Reference -like '*#REF!*'
Open-OfficeExcel -Path .\Finance\Budget-2027.xlsx {
    foreach ($name in $broken) {
        Remove-OfficeExcelNamedRange -Name $name.Name
    }
}
```

Removes every defined name that points to #REF! and saves the workbook.

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
