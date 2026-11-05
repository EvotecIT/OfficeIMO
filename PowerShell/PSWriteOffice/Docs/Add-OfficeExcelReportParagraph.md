---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Add-OfficeExcelReportParagraph
## SYNOPSIS
Adds a paragraph of text to the report.

## SYNTAX
### Default (Default)
```powershell
Add-OfficeExcelReportParagraph [-Text] <string> [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Adds a paragraph of text to the report.

## EXAMPLES

### EXAMPLE 1
```powershell
New-OfficeExcel -Path .\Reports\Operations.xlsx {
    ExcelReportSheet 'Summary' {
        ExcelReportParagraph 'Figures cover all production services from 1 to 31 October.'
    }
}
```

Adds a paragraph under the title.

### EXAMPLE 2
```powershell
New-OfficeExcel -Path .\Reports\Service-Review-October.xlsx {
    ExcelReportSheet 'Review' {
        ExcelReportHeading 'Availability by site'
        ExcelReportParagraph 'Availability is measured every minute from two external probes. Planned maintenance announced five days ahead is excluded.'
        ExcelReportTable (Import-OfficeCsv -Path .\Data\Site-Availability.csv)
    }
}
```

Describes how availability is measured before the site table, so readers do not need a separate methods document.

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
