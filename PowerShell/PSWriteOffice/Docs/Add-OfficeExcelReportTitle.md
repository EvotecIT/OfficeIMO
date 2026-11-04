---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Add-OfficeExcelReportTitle
## SYNOPSIS
Adds the report title and subtitle.

## SYNTAX
### Default (Default)
```powershell
Add-OfficeExcelReportTitle [-Title] <string> [[-Subtitle] <string>] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Adds the report title and subtitle.

## EXAMPLES

### EXAMPLE 1
```powershell
New-OfficeExcel -Path .\Reports\Operations.xlsx {
    ExcelReportSheet 'Summary' {
        ExcelReportTitle 'Operations summary'
    }
}
```

Adds the title block at the top of a report sheet.

### EXAMPLE 2
```powershell
$period = (Get-Date).ToString('MMMM yyyy')
New-OfficeExcel -Path .\Reports\Capacity-Review.xlsx {
    ExcelReportSheet 'Capacity' {
        ExcelReportTitle 'Capacity review' "Storage and compute, $period"
        ExcelReportTable (Import-OfficeCsv -Path .\Data\Capacity.csv) 'Clusters'
    }
}
```

Adds a title and a subtitle built from the current month, so a scheduled run labels itself.

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

### -Subtitle
Dashboard subtitle.

```yaml
Type: String
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: 1
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Title
Chart title.

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

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `None`

## OUTPUTS

- `None`

## RELATED LINKS

- None
