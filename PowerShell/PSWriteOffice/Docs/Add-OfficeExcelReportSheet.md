---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Add-OfficeExcelReportSheet
## SYNOPSIS
Creates a worksheet through the OfficeIMO sheet composer and runs report-block cmdlets inside it.

## SYNTAX
### __AllParameterSets
```powershell
Add-OfficeExcelReportSheet [-Name] <string> [[-Content] <scriptblock>] [-SectionHeaderFillColor <string>] [-KeyFillColor <string>] [-NoAutoFit] [-AutoFitRows] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Creates a worksheet through the OfficeIMO sheet composer and runs report-block cmdlets inside it.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> New-OfficeExcel -Path .\report.xlsx {
  Add-OfficeExcelReportSheet -Name Summary {
    Add-OfficeExcelReportTitle -Title 'Operational Summary' -Subtitle 'Current view'
    Add-OfficeExcelReportKpiRow -InputObject @{ Ready = 12; Blocked = 2 }
  }
}
```

Creates a report-oriented worksheet with title and KPI blocks.

## PARAMETERS

### -AutoFitRows
Auto-fit row heights during composer finalization.

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

### -Content
Report block script to run inside the composer context.

```yaml
Type: ScriptBlock
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: 1
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -KeyFillColor
Override the key-cell fill color used by KPI and property blocks.

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

### -Name
Name of the report worksheet to create.

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

### -NoAutoFit
Skip composer auto-fit finalization.

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

### -PassThru
Emit the created worksheet.

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

### -SectionHeaderFillColor
Override the section-header fill color.

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

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `None`

## OUTPUTS

- `None`

## RELATED LINKS

- None
