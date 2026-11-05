---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Set-OfficeExcelDocumentProperty
## SYNOPSIS
Sets a built-in or application document property on an Excel workbook.

## SYNTAX
### Default (Default)
```powershell
Set-OfficeExcelDocumentProperty [-Name] <string> [[-Value] <Object>] [-Document <ExcelDocument>] [-PassThru] [-Custom] [<CommonParameters>]
```

## DESCRIPTION
Sets a built-in or application document property on an Excel workbook.

## EXAMPLES

### EXAMPLE 1
```powershell
Open-OfficeExcel -Path .\Reports\Q3-Sales.xlsx {
    Set-OfficeExcelDocumentProperty Title 'Q3 sales'
}
```

Sets the built-in Title property.

### EXAMPLE 2
```powershell
Open-OfficeExcel -Path .\Reports\Q3-Sales.xlsx {
    Set-OfficeExcelDocumentProperty -Name Owner -Value 'Sales Operations' -Custom
    Set-OfficeExcelDocumentProperty -Name Classification -Value 'Internal' -Custom
    Set-OfficeExcelDocumentProperty -Name ReviewDate -Value (Get-Date).AddMonths(6) -Custom
}
```

Adds owner, classification and review-date custom properties that document management can search.

## PARAMETERS

### -Custom
Treat the property as a custom workbook property.

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
Workbook to update when provided explicitly.

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

### -Name
Property name to update.

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
Emit the updated workbook.

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

### -Value
Property value.

```yaml
Type: Object
Parameter Sets: Default
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
