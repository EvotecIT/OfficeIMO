---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Close-OfficeVisio
## SYNOPSIS
Releases a live Visio document without saving.

## SYNTAX
### Document (Default)
```powershell
Close-OfficeVisio [-Document] <VisioDocument> [<CommonParameters>]
```

## DESCRIPTION
Releases a live Visio document without saving.

## EXAMPLES

### EXAMPLE 1
```powershell
Close-OfficeVisio -Document $document
```


## PARAMETERS

### -Document
Live document to release.

```yaml
Type: VisioDocument
Parameter Sets: Document
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `OfficeIMO.Visio.VisioDocument`

## OUTPUTS

- `System.Void`

## RELATED LINKS

- None
