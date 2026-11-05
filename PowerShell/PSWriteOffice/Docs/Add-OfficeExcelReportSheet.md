---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Add-OfficeExcelReportSheet
## SYNOPSIS
Creates a report worksheet and lays out report blocks inside it.

## SYNTAX
### Default (Default)
```powershell
Add-OfficeExcelReportSheet [-Name] <string> [[-Content] <scriptblock>] [-Document <ExcelDocument>] [-HeadingFillColor <OfficeColor>] [-KeyFillColor <OfficeColor>] [-NoAutoFit] [-AutoFitRows] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Creates a report worksheet and lays out report blocks inside it.

## EXAMPLES

### EXAMPLE 1
```powershell
New-OfficeExcel -Path .\Reports\Operations.xlsx {
    ExcelReportSheet 'Summary' {
        ExcelReportTitle 'Operations summary' 'Week 43, 2026'
        ExcelReportKpi ([ordered]@{ Uptime = '99.97%'; Incidents = 4; 'Changes shipped' = 37 })
    }
}
```

Adds a Summary report sheet with a title and a KPI row, laid out by the report builder in the workbook theme.

### EXAMPLE 2
```powershell
$sites = Import-OfficeCsv -Path .\Data\Site-Availability.csv
$legend = @(
    [pscustomobject]@{ Status = 'Met'; Meaning = 'Availability at or above the 99.9% target' }
    [pscustomobject]@{ Status = 'Missed'; Meaning = 'Below target; a service credit may apply' }
)
New-OfficeExcel -Path .\Reports\Service-Review-October.xlsx -Theme Boardroom {
    ExcelReportSheet 'Review' {
        ExcelReportTitle 'Service review' 'October 2026'
        ExcelReportKpi ([ordered]@{ Availability = '99.94%'; 'P1 incidents' = 1; 'Mean time to restore' = '42 min'; 'Changes' = 128 })
        ExcelReportCallout Warning 'Two sites missed target' 'Gdansk and Porto were below 99.9% after the network maintenance on 14 October.'
        ExcelReportHeading 'Availability by site'
        ExcelReportTable $sites 'Sites'
        ExcelReportChart -Type Column -Category Site -Value Availability -Title 'Availability by site'
        ExcelReportSpacer
        ExcelReportLegend 'Status' -Header Status, Meaning -InputObject $legend
    }
}
```

Lays out a complete service review: title, KPIs, a callout, a site table with its chart and a legend, all styled by the theme.

## PARAMETERS

### -AutoFitRows
Auto-fit row heights during composer finalization.

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

### -Content
Excel child commands executed in the current content scope.

```yaml
Type: ScriptBlock
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: 1
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

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

### -HeadingFillColor
Fill colour for report section headings.

```yaml
Type: OfficeColor
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -KeyFillColor
Override the key-cell fill color used by KPI and property blocks.

```yaml
Type: OfficeColor
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
Name of the report worksheet to create.

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

### -NoAutoFit
Retains native column widths without fitting them to the data.

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

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `OfficeIMO.Excel.ExcelDocument`

## OUTPUTS

- `None`

## RELATED LINKS

- None
