---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Add-OfficeExcelReportTitle
## SYNOPSIS
Adds a title block to the current Excel report sheet.

## SYNTAX
### __AllParameterSets
```powershell
Add-OfficeExcelReportTitle [-Title] <string> [[-Subtitle] <string>] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Adds a title block to the current Excel report sheet.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> New-OfficeExcel -Path .\Operations.xlsx {
    Add-OfficeExcelReportSheet -Name Summary {
        Add-OfficeExcelReportTitle -Title 'Operational Summary' -Subtitle 'Current month'
        Add-OfficeExcelReportKpiRow -InputObject @{ Revenue = 125000; Incidents = 3; Status = 'Ready' }
    }
}
```

Uses the OfficeIMO sheet composer through PSWriteOffice's thin report-block wrapper.

## PARAMETERS

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

### -Subtitle
Optional subtitle text.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: 1
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Title
Title text.

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

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `None`

## OUTPUTS

- `None`

## RELATED LINKS

- None
