---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# New-OfficeWordMarkdownOptions
## SYNOPSIS
Creates native Word-to-Markdown options for Convert-OfficeDocument.

## SYNTAX
### Default (Default)
```powershell
New-OfficeWordMarkdownOptions [-Font <string>] [-EnableUnderline] [-EnableHighlight] [-ImageExportMode <WordMarkdownImageExportMode>] [-ImageDirectory <string>] [<CommonParameters>]
```

## DESCRIPTION
Creates native Word-to-Markdown options for Convert-OfficeDocument.

## EXAMPLES

### EXAMPLE 1
```powershell
$options = New-OfficeWordMarkdownOptions -EnableUnderline -EnableHighlight
Convert-OfficeDocument -Path ./Runbook.docx -To Markdown -OutputPath ./Runbook.md -Options $options
```


## PARAMETERS

### -EnableHighlight
Wrap highlighted text in double equals signs.

```yaml
Type: SwitchParameter
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -EnableUnderline
Wrap underlined text in HTML underline tags.

```yaml
Type: SwitchParameter
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Font
Font family whose runs become inline code; omission uses the native monospace resolver.

```yaml
Type: String
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ImageDirectory
Filesystem folder for images when File mode is selected; this command does not create it.

```yaml
Type: String
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ImageExportMode
Embed images as Base64 or save them as File references.

```yaml
Type: WordMarkdownImageExportMode
Parameter Sets: Default
Aliases: None
Possible values: Base64, File

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

- `OfficeIMO.Word.Markdown.WordMarkdownExportOptions`

## RELATED LINKS

- None
