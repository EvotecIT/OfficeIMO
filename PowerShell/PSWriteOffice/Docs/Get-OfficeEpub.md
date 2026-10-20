---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Get-OfficeEpub
## SYNOPSIS
Loads a publication for package-preserving inspection or editing.

## SYNTAX
### __AllParameterSets
```powershell
Get-OfficeEpub [-Path] <string> [<CommonParameters>]
```

## DESCRIPTION
Loads a publication for package-preserving inspection or editing.

## EXAMPLES

### EXAMPLE 1
```powershell
Get-Item ./book.epub | Get-OfficeEpub
```


## PARAMETERS

### -Path
Literal EPUB path.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: FullName, FilePath, LiteralPath
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: True (ByValue, ByPropertyName)
Accept wildcard characters: False
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `System.String`

## OUTPUTS

- `OfficeIMO.Epub.EpubPublication`

## RELATED LINKS

- None
