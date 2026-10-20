---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# New-OfficeEpub
## SYNOPSIS
Creates an EPUB publication with native navigation and metadata.

## SYNTAX
### __AllParameterSets
```powershell
New-OfficeEpub [-Title] <string> [-Language <string>] [-Identifier <string>] [-Author <string>] [-Version <EpubVersion>] [<CommonParameters>]
```

## DESCRIPTION
Creates an EPUB publication with native navigation and metadata.

## EXAMPLES

### EXAMPLE 1
```powershell
$book = New-OfficeEpub -Title Operations -Language en -Author Example
```


## PARAMETERS

### -Author
Primary author.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Identifier
Optional stable publication identifier.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Language
BCP 47 language.

```yaml
Type: String
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
Publication title.

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

### -Version
EPUB package generation.

```yaml
Type: EpubVersion
Parameter Sets: __AllParameterSets
Aliases: None
Possible values: Epub2, Epub3

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

- `OfficeIMO.Epub.EpubPublication`

## RELATED LINKS

- None
