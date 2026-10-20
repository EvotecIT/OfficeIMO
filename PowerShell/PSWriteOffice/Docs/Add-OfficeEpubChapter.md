---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Add-OfficeEpubChapter
## SYNOPSIS
Adds chapter content, reading order and a navigation entry.

## SYNTAX
### Text (Default)
```powershell
Add-OfficeEpubChapter [-Title] <string> -Publication <EpubPublication> -Text <string> [-Id <string>] [-ContainerPath <string>] [-PassThru] [<CommonParameters>]
```

### Xhtml
```powershell
Add-OfficeEpubChapter [-Title] <string> -Publication <EpubPublication> -Xhtml <string> [-Id <string>] [-ContainerPath <string>] [-PassThru] [<CommonParameters>]
```

### Markdown
```powershell
Add-OfficeEpubChapter [-Title] <string> -Publication <EpubPublication> -Markdown <string> [-Id <string>] [-ContainerPath <string>] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Adds chapter content, reading order and a navigation entry.

## EXAMPLES

### EXAMPLE 1
```powershell
$book | Add-OfficeEpubChapter -Title Opening -Text "Welcome to this handbook."
```


## PARAMETERS

### -ContainerPath
Optional package-relative XHTML path.

```yaml
Type: String
Parameter Sets: Text, Xhtml, Markdown
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Id
Optional stable manifest identifier; omission chooses an unused chapter number.

```yaml
Type: String
Parameter Sets: Text, Xhtml, Markdown
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Markdown
Markdown rendered by the existing Markdown engine. Unsupported raw HTML remains subject to EPUB validation.

```yaml
Type: String
Parameter Sets: Markdown
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PassThru
Return the created or changed object.

```yaml
Type: SwitchParameter
Parameter Sets: Text, Xhtml, Markdown
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Publication
Live editable publication.

```yaml
Type: EpubPublication
Parameter Sets: Text, Xhtml, Markdown
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -Text
Plain text, escaped into XHTML paragraphs.

```yaml
Type: String
Parameter Sets: Text
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Title
Chapter title.

```yaml
Type: String
Parameter Sets: Text, Xhtml, Markdown
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Xhtml
Well-formed XHTML body fragment.

```yaml
Type: String
Parameter Sets: Xhtml
Aliases: None
Possible values:

Required: True
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

- `OfficeIMO.Epub.EpubManifestItem`

## RELATED LINKS

- None
