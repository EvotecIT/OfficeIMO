---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Add-OfficeExcelReportSpacer
## SYNOPSIS
Adds vertical space to the report.

## SYNTAX
### Default (Default)
```powershell
Add-OfficeExcelReportSpacer [[-Rows] <int>] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Adds vertical space to the report.

## EXAMPLES

### EXAMPLE 1
```powershell
New-OfficeExcel -Path .\Reports\Operations.xlsx {
    ExcelReportSheet 'Summary' {
        ExcelReportTitle 'Operations summary'
        ExcelReportSpacer
        ExcelReportParagraph 'Week 43, 2026'
    }
}
```

Leaves one empty row between two report blocks.

### EXAMPLE 2
```powershell
New-OfficeExcel -Path .\Reports\Helpdesk-Review.xlsx {
    ExcelReportSheet 'Review' {
        ExcelReportKpi ([ordered]@{ Opened = 412; Closed = 398; 'Median age' = '1.6 days' })
        ExcelReportSpacer 2
        ExcelReportTable (Import-OfficeCsv -Path .\Data\Tickets-By-Team.csv) 'Tickets by team'
    }
}
```

Leaves two rows between the KPI cards and the detail table to give the summary room.

## PARAMETERS

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

### -Rows
Auto-fit all rows.

```yaml
Type: Int32
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
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
