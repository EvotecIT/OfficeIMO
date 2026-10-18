---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Save-OfficeMarkdown
## SYNOPSIS
Saves a Markdown document without changing its lifetime.

## SYNTAX
### __AllParameterSets
```powershell
Save-OfficeMarkdown [-Document] <MarkdownDoc> [-Path] <string> [-WriteOptions <MarkdownWriteOptions>] [-WriteProfile <OfficeMarkdownWriteProfile>] [-ImageRenderingMode <MarkdownImageRenderingMode>] [-LineEnding <string>] [-UnorderedListMarker <string>] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Saves a Markdown document without changing its lifetime.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $doc | Save-OfficeMarkdown -Path .\Report.md
```

Writes the Markdown artifact and keeps the document available for further changes.

## PARAMETERS

### -Document
Markdown document to save.

```yaml
Type: MarkdownDoc
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -ImageRenderingMode
Controls how Markdown images are serialized.

```yaml
Type: MarkdownImageRenderingMode
Parameter Sets: __AllParameterSets
Aliases: None
Possible values: RichMarkdown, PortableMarkdown, Html

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -LineEnding
Markdown line ending: CRLF, LF, CR, or a literal line ending string.

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
Emit the Markdown document rather than the saved file.

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

### -Path
Destination Markdown path.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: FilePath
Possible values:

Required: True
Position: 1
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -UnorderedListMarker
Unordered list marker: '-', '*', or '+'.

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

### -WriteOptions
Optional Markdown writer options.

```yaml
Type: MarkdownWriteOptions
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -WriteProfile
Friendly Markdown writer profile.

```yaml
Type: OfficeMarkdownWriteProfile
Parameter Sets: __AllParameterSets
Aliases: None
Possible values: OfficeIMO, Portable, HtmlImage

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `OfficeIMO.Markdown.MarkdownDoc`

## OUTPUTS

- `OfficeIMO.Markdown.MarkdownDoc`

## RELATED LINKS

- None
