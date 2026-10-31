---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Get-OfficeEpubChapter
## SYNOPSIS
Projects chapters to searchable PowerShell objects.

## SYNTAX
### __AllParameterSets
```powershell
Get-OfficeEpubChapter [-Document <EpubPublication>] [<CommonParameters>]
```

## DESCRIPTION
Projects chapters to searchable PowerShell objects.

## EXAMPLES

### EXAMPLE 1
```powershell
$book | Get-OfficeEpubChapter | Select-Object Order, Title, Text
```


## PARAMETERS

### -Document
Gets or sets an explicit live document; omission uses the enclosing content block.

```yaml
Type: EpubPublication
Parameter Sets: __AllParameterSets
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

- `OfficeIMO.Epub.EpubPublication`

## OUTPUTS

- `OfficeIMO.Epub.EpubChapter`

## RELATED LINKS

- None
