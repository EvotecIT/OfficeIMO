---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Set-OfficeEpubMetadata
## SYNOPSIS
Updates the primary title, language or creator through the package model.

## SYNTAX
### __AllParameterSets
```powershell
Set-OfficeEpubMetadata [-Title <string>] [-Language <string>] [-Author <string>] [-PassThru] [-Document <EpubPublication>] [<CommonParameters>]
```

## DESCRIPTION
Updates the primary title, language or creator through the package model.

## EXAMPLES

### EXAMPLE 1
```powershell
$book | Set-OfficeEpubMetadata -Title "Operations handbook"
```


## PARAMETERS

### -Author
New primary author.

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

### -Language
New BCP 47 language.

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

### -PassThru
Return the created or changed object.

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

### -Title
New primary title.

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

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `OfficeIMO.Epub.EpubPublication`

## OUTPUTS

- `OfficeIMO.Epub.EpubPublication`

## RELATED LINKS

- None
