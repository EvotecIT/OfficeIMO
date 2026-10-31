---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Close-OfficeEpub
## SYNOPSIS
Completes an EPUB publication's lifecycle without saving.

## SYNTAX
### __AllParameterSets
```powershell
Close-OfficeEpub -Document <EpubPublication> [<CommonParameters>]
```

## DESCRIPTION
Completes an EPUB publication's lifecycle without saving.

## EXAMPLES

### EXAMPLE 1
```powershell
Close-OfficeEpub -Document $book
```


## PARAMETERS

### -Document
Live publication to close without saving.

```yaml
Type: EpubPublication
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `OfficeIMO.Epub.EpubPublication`

## OUTPUTS

- `None`

## RELATED LINKS

- None
