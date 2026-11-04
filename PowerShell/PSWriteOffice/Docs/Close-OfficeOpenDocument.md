---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Close-OfficeOpenDocument
## SYNOPSIS
Releases a live OpenDocument document without saving.

## SYNTAX
### Document (Default)
```powershell
Close-OfficeOpenDocument -Document <OdfDocument> [<CommonParameters>]
```

## DESCRIPTION
Releases a live OpenDocument document without saving.

## EXAMPLES

### EXAMPLE 1
```powershell
Close-OfficeOpenDocument -Document $document
```


## PARAMETERS

### -Document
Live document to release.

```yaml
Type: OdfDocument
Parameter Sets: Document
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

- `OfficeIMO.OpenDocument.OdfDocument`

## OUTPUTS

- `System.Void`

## RELATED LINKS

- None
