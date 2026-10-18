---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Add-OfficeExcelReportKpiRow
## SYNOPSIS
Adds a KPI row to the current Excel report sheet.

## SYNTAX
### __AllParameterSets
```powershell
Add-OfficeExcelReportKpiRow [-InputObject] <Object> [-PerRow <int>] [-LabelFillColor <string>] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Adds a KPI row to the current Excel report sheet.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> New-OfficeExcel -Path .\Operations.xlsx {
    Add-OfficeExcelReportSheet -Name Summary {
        Add-OfficeExcelReportKpiRow -InputObject @{ Revenue = 125000; Incidents = 3; Status = 'Ready' } -PerRow 3
    }
}
```

Renders PowerShell key/value data as a KPI row through the OfficeIMO sheet composer.

## PARAMETERS

### -InputObject
Hashtable or objects with Label/Value, Key/Value, Name/Value, or Title/Value properties.

```yaml
Type: Object
Parameter Sets: __AllParameterSets
Aliases: Data
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -LabelFillColor
Optional fill color for KPI labels.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PassThru
Emit the object created or changed by the command.

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

### -PerRow
Number of KPI cards per rendered row.

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

- `None`

## RELATED LINKS

- None
