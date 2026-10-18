---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Set-OfficeExcelDateSystem
## SYNOPSIS
Sets the workbook date system used for numeric date serials.

## SYNTAX
### __AllParameterSets
```powershell
Set-OfficeExcelDateSystem [-Document] <ExcelDocument> [-DateSystem] <string> [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Sets the workbook date system used for numeric date serials.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $workbook | Set-OfficeExcelDateSystem -DateSystem 1904 -PassThru | Save-OfficeExcel
```

Marks the workbook to use Excel's 1904 date system before saving.

## PARAMETERS

### -DateSystem
Date system to use for Excel date serials.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: None
Possible values: 1900, 1904, NineteenHundred, NineteenFour

Required: True
Position: 1
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Document
Workbook to update.

```yaml
Type: ExcelDocument
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -PassThru
Emit the workbook for further pipeline operations.

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

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `OfficeIMO.Excel.ExcelDocument`

## OUTPUTS

- `OfficeIMO.Excel.ExcelDocument`

## RELATED LINKS

- None
