---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Add-OfficeOpenDocumentSlide
## SYNOPSIS
Adds a slide to an OpenDocument presentation and optionally runs nested slide content.

## SYNTAX
### __AllParameterSets
```powershell
Add-OfficeOpenDocumentSlide [[-Name] <string>] [[-Content] <scriptblock>] [-Document <OdpPresentation>] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Adds a slide to an OpenDocument presentation and optionally runs nested slide content.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> Add-OfficeOpenDocumentSlide -Name 'Summary' -Content {
    Add-OfficeOpenDocumentTextBox -Text 'Quarterly summary' -X 2 -Y 2 -Width 20 -Height 3
}
```


## PARAMETERS

### -Content
Nested slide commands that use this slide as their current target.

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
OpenDocument presentation. Omit inside New-OfficeOpenDocument -Content.

```yaml
Type: OdpPresentation
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
Optional unique slide name.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PassThru
Emit the created slide.

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

- `OfficeIMO.OpenDocument.OdpPresentation`

## OUTPUTS

- `OfficeIMO.OpenDocument.OdpSlide`

## RELATED LINKS

- None
