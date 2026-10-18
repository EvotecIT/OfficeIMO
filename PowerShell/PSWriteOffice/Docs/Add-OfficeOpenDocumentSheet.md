---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Add-OfficeOpenDocumentSheet
## SYNOPSIS
Adds a worksheet to an OpenDocument spreadsheet and optionally runs nested cell content.

## SYNTAX
### __AllParameterSets
```powershell
Add-OfficeOpenDocumentSheet [-Name] <string> [[-Content] <scriptblock>] [-Document <OdsDocument>] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Adds a worksheet to an OpenDocument spreadsheet and optionally runs nested cell content.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> Add-OfficeOpenDocumentSheet -Name 'Data' -Content {
    Set-OfficeOpenDocumentCell -Row 0 -Column 0 -Value 'Status'
}
```


## PARAMETERS

### -Content
Nested cell commands that use this worksheet as their current target.

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

### -Document
OpenDocument spreadsheet. Omit inside New-OfficeOpenDocument -Content.

```yaml
Type: OdsDocument
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -Name
Worksheet name.

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

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `OfficeIMO.OpenDocument.OdsDocument`

## OUTPUTS

- `OfficeIMO.OpenDocument.OdsSheet`

## RELATED LINKS

- None
