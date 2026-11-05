---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Add-OfficeExcelReportHeading
## SYNOPSIS
Adds a section heading to the report.

## SYNTAX
### Default (Default)
```powershell
Add-OfficeExcelReportHeading [-Text] <string> [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Adds a section heading to the report.

## EXAMPLES

### EXAMPLE 1
```powershell
New-OfficeExcel -Path .\Reports\Operations.xlsx {
    ExcelReportSheet 'Summary' {
        ExcelReportHeading 'Incidents'
    }
}
```

Starts a new section of the report.

### EXAMPLE 2
```powershell
New-OfficeExcel -Path .\Reports\Security-Review.xlsx {
    ExcelReportSheet 'Review' {
        ExcelReportTitle 'Security review' 'Q3 2026'
        ExcelReportHeading 'Privileged accounts'
        ExcelReportTable (Import-OfficeCsv -Path .\Data\Privileged-Accounts.csv)
        ExcelReportHeading 'Firewall changes'
        ExcelReportTable (Import-OfficeCsv -Path .\Data\Firewall-Changes.csv)
    }
}
```

Gives each part of a security report its own heading and table.

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

### -Text
Comment text.

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
