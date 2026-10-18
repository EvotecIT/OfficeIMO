---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Add-OfficeExcelReportCallout
## SYNOPSIS
Adds a colored callout block to the current Excel report sheet.

## SYNTAX
### __AllParameterSets
```powershell
Add-OfficeExcelReportCallout [[-Kind] <OfficeExcelReportCalloutKind>] [-Title] <string> [-Body] <string> [-WidthColumns <int>] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Adds a colored callout block to the current Excel report sheet.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> New-OfficeExcel -Path .\Operations.xlsx {
    Add-OfficeExcelReportSheet -Name Summary {
        Add-OfficeExcelReportCallout -Kind Warning -Title 'Manual validation' -Body 'Open the workbook in desktop Excel before publishing pivot-heavy reports.'
    }
}
```

Renders a composer callout block using the current report theme.

## PARAMETERS

### -Body
Callout body text.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: True
Position: 2
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Kind
Callout kind. Supported values include info, success, warning, error, and critical.

```yaml
Type: OfficeExcelReportCalloutKind
Parameter Sets: __AllParameterSets
Aliases: None
Possible values: Info, Success, Warning, Error, Critical

Required: False
Position: 0
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

### -Title
Callout title.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: True
Position: 1
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -WidthColumns
Width of the highlighted callout band in worksheet columns.

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
