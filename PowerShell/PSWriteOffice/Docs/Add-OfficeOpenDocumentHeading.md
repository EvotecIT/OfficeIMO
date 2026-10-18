---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Add-OfficeOpenDocumentHeading
## SYNOPSIS
Adds a heading to an OpenDocument text document.

## SYNTAX
### __AllParameterSets
```powershell
Add-OfficeOpenDocumentHeading [-Text] <string> [-Document <OdtDocument>] [-Level <int>] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Adds a heading to an OpenDocument text document.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> Add-OfficeOpenDocumentHeading -Text 'Results' -Level 2
```


## PARAMETERS

### -Document
OpenDocument text document. Omit inside New-OfficeOpenDocument -Content.

```yaml
Type: OdtDocument
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -Level
Heading level from 1 through 10.

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

### -PassThru
Emit the created heading paragraph.

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

### -Text
Heading text.

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

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `OfficeIMO.OpenDocument.OdtDocument`

## OUTPUTS

- `OfficeIMO.OpenDocument.OdtParagraph`

## RELATED LINKS

- None
