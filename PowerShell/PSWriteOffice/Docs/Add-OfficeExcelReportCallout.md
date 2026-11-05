---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Add-OfficeExcelReportCallout
## SYNOPSIS
Adds a highlighted info, success, warning or error callout.

## SYNTAX
### Default (Default)
```powershell
Add-OfficeExcelReportCallout [[-Kind] <OfficeExcelReportCalloutKind>] [-Title] <string> [-Text] <string> [-Span <int>] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Adds a highlighted info, success, warning or error callout.

## EXAMPLES

### EXAMPLE 1
```powershell
New-OfficeExcel -Path .\Reports\Operations.xlsx {
    ExcelReportSheet 'Summary' {
        ExcelReportCallout Info 'Data source' 'Monitoring export from 1 November, 06:00.'
    }
}
```

Highlights a note for readers in an information callout.

### EXAMPLE 2
```powershell
$failed = Import-OfficeCsv -Path .\Data\Backup-Results.csv | Where-Object Result -ne 'Success'
New-OfficeExcel -Path .\Reports\Backup-Check.xlsx {
    ExcelReportSheet 'Backups' {
        ExcelReportTitle 'Nightly backup check' (Get-Date).ToString('d MMMM yyyy')
        if ($failed) {
            ExcelReportCallout Error 'Backups failed' "$(@($failed).Count) jobs did not complete. Details are below."
            ExcelReportTable $failed 'Failed jobs'
        } else {
            ExcelReportCallout Success 'All backups completed' 'Every scheduled job finished successfully.'
        }
    }
}
```

Adds a success or error callout depending on whether the backup checks passed, so the outcome is the first thing a reader sees.

## PARAMETERS

### -Kind
Metadata kind to add.

```yaml
Type: OfficeExcelReportCalloutKind
Parameter Sets: Default
Aliases: None
Possible values: Info, Success, Warning, Error, Critical

Required: False
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

### -Span
Number of worksheet columns occupied by the report block.

```yaml
Type: Int32
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Text
Comment text.

```yaml
Type: String
Parameter Sets: Default
Aliases: None
Possible values:

Required: True
Position: 2
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
Position: 1
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
