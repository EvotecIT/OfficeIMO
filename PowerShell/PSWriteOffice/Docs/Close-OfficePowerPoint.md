---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Close-OfficePowerPoint
## SYNOPSIS
Closes a live presentation without saving it.

## SYNTAX
### Document (Default)
```powershell
Close-OfficePowerPoint [-Document <PowerPointDocument>] [<CommonParameters>]
```

## DESCRIPTION
Releases the native document and writes no output. Save changes explicitly with Save-OfficePowerPoint before closing. Documents configured for SaveOnDispose must be reopened with explicit persistence.

## EXAMPLES

### EXAMPLE 1
```powershell
$deck = Open-OfficePowerPoint -Path ./Review.pptx
Save-OfficePowerPoint -Document $deck
Close-OfficePowerPoint -Document $deck
```


## PARAMETERS

### -Document
Gets or sets an explicit live document; omission uses the enclosing content block.

```yaml
Type: PowerPointDocument
Parameter Sets: Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `OfficeIMO.PowerPoint.PowerPointDocument`

## OUTPUTS

- `None`

## RELATED LINKS

- None
