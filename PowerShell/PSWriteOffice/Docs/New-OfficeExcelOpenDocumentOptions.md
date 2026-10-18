---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# New-OfficeExcelOpenDocumentOptions
## SYNOPSIS
Creates Excel/OpenDocument conversion settings.

## SYNTAX
### __AllParameterSets
```powershell
New-OfficeExcelOpenDocumentOptions [-LossPolicy <OdfConversionLossPolicy>] [-IncludeBasicStyles] [-MaximumExpandedCells <Int64>] [-MaximumRows <Int32>] [-MaximumColumns <Int32>] [<CommonParameters>]
```

## DESCRIPTION
Creates Excel/OpenDocument conversion settings.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $options = New-OfficeExcelOpenDocumentOptions -IncludeBasicStyles -MaximumRows 10000 -MaximumColumns 100
ConvertTo-OfficeOpenDocument -Path .\Data.xlsx -OutputPath .\Data.ods -ExcelOptions $options
```


## PARAMETERS

### -IncludeBasicStyles
Copy common font, fill, and number-format styles.

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

### -LossPolicy
Whether conversion loss is reported or rejected.

```yaml
Type: OdfConversionLossPolicy
Parameter Sets: __AllParameterSets
Aliases: None
Possible values: ReportOnly, ThrowOnSkippedOrUnsupported, ThrowOnAnyLoss

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -MaximumColumns
Maximum spreadsheet columns.

```yaml
Type: Int32
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -MaximumExpandedCells
Maximum cells materialized during conversion.

```yaml
Type: Int64
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -MaximumRows
Maximum spreadsheet rows.

```yaml
Type: Int32
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

- `None`

## OUTPUTS

- `OfficeIMO.Excel.OpenDocument.ExcelOpenDocumentConversionOptions`

## RELATED LINKS

- None
