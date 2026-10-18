---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Add-OfficeExcelReportSpacer
## SYNOPSIS
Adds vertical spacing to the current Excel report sheet.

## SYNTAX
### __AllParameterSets
```powershell
Add-OfficeExcelReportSpacer [[-Rows] <int>] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Adds vertical spacing to the current Excel report sheet.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> New-OfficeExcel -Path .\Operations.xlsx {
    Add-OfficeExcelReportSheet -Name Summary {
        Add-OfficeExcelReportTitle -Title 'Operational Summary'
        Add-OfficeExcelReportSpacer -Rows 2
        Add-OfficeExcelReportSection -Text 'Details'
    }
}
```

Advances the composer cursor before adding the next report block.

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

### -Rows
Rows to advance. Defaults to the composer theme spacing.

```yaml
Type: Int32
Parameter Sets: __AllParameterSets
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
