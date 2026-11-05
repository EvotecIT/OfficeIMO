---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Set-OfficeExcelNamedRange
## SYNOPSIS
Creates or updates a defined name.

## SYNTAX
### Default (Default)
```powershell
Set-OfficeExcelNamedRange [-Name] <string> [-Range] <string> [-Document <ExcelDocument>] [-Sheet <OfficeSheetSelector>] [-Hidden] [-ValidationMode <ExcelDefinedNameValidationMode>] [-Global] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Creates or updates a defined name.

## EXAMPLES

### EXAMPLE 1
```powershell
Open-OfficeExcel -Path .\Finance\Budget-2027.xlsx {
    ExcelSheet 'Assumptions' {
        ExcelNamedRange TaxRate 'B2'
    }
}
```

Names B2 on the Assumptions sheet TaxRate so formulas can use =Amount*TaxRate.

### EXAMPLE 2
```powershell
New-OfficeExcel -Path .\Finance\Pricing-Model.xlsx {
    ExcelSheet 'Inputs' {
        ExcelCell A1 'Exchange rate EUR/USD'
        ExcelCell B1 1.08
        ExcelCell A2 'Target margin'
        ExcelCell B2 0.32 -NumberFormat '0%'
        ExcelNamedRange EurUsd 'B1'
        ExcelNamedRange TargetMargin 'B2'
    }
    ExcelSheet 'Prices' {
        ExcelCell A1 'List price (USD)'
        ExcelCell B1 -Formula '=1000*EurUsd/(1-TargetMargin)'
    }
}
```

Writes the model inputs and names each one, so the calculation sheets read like the business rules they implement.

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
Force a workbook-global name even inside a sheet block.

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

### -Hidden
Mark the defined name as hidden.

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
Name of the defined range.

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

### -Range
A1 cell range on the selected worksheet.

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
Validate or sanitize the defined name.

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
