---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Set-OfficeExcelDocumentProperty
## SYNOPSIS
Sets a built-in or application document property on an Excel workbook.

## SYNTAX
### __AllParameterSets
```powershell
Set-OfficeExcelDocumentProperty [-Name] <string> [[-Value] <Object>] [-Document <ExcelDocument>] [-PassThru] [-Custom] [<CommonParameters>]
```

## DESCRIPTION
Sets a built-in or application document property on an Excel workbook.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> New-OfficeExcel -Path .\Report.xlsx {
    Set-OfficeExcelDocumentProperty -Name Title -Value 'Operational dashboard'
    Set-OfficeExcelDocumentProperty -Name Department -Value 'Operations' -Custom
}
```

Updates built-in or custom workbook properties through the current Excel DSL context.

## PARAMETERS

### -Custom
Treat the property as a custom workbook property.

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

### -Document
Workbook to update when provided explicitly.

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

### -Name
Property name to update.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PassThru
Emit the updated workbook.

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

### -Value
Property value.

```yaml
Type: Object
Parameter Sets: __AllParameterSets
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

- `OfficeIMO.Excel.ExcelDocument`

## RELATED LINKS

- None
