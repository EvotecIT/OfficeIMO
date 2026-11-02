---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Save-OfficeMarkdown
## SYNOPSIS
Saves a live Markdown document using its native save policy and keeps it available.

## SYNTAX
### Document (Default)
```powershell
Save-OfficeMarkdown [[-Path] <string>] -Document <MarkdownDocument> [-WriteOptions <MarkdownSaveOptions>] [-WriteProfile <OfficeMarkdownWriteProfile>] [-ImageRenderingMode <MarkdownImageRenderingMode>] [-LineEnding <OfficeLineEnding>] [-UnorderedListMarker <string>] [-AllowLoss] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Saves a live Markdown document using its native save policy and keeps it available.

## EXAMPLES

### EXAMPLE 1
```powershell
$document | Save-OfficeMarkdown -PassThru
```


### EXAMPLE 2
```powershell
Save-OfficeMarkdown -Document $document -Path ./Publish/Report.md -LineEnding CRLF -PassThru
```


## PARAMETERS

### -AllowLoss
Accept explicit native-save omissions, overriding WriteOptions when supplied.

```yaml
Type: SwitchParameter
Parameter Sets: Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Document
Markdown document to save.

```yaml
Type: MarkdownDocument
Parameter Sets: Document
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -ImageRenderingMode
Controls Markdown image serialization.

```yaml
Type: MarkdownImageRenderingMode
Parameter Sets: Document
Aliases: None
Possible values: RichMarkdown, PortableMarkdown, Html

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -LineEnding
Output line endings.

```yaml
Type: OfficeLineEnding
Parameter Sets: Document
Aliases: None
Possible values: LF, CRLF, CR

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PassThru
Return the saved file.

```yaml
Type: SwitchParameter
Parameter Sets: Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Path
Optional destination; omission saves to the associated path.

```yaml
Type: String
Parameter Sets: Document
Aliases: None
Possible values:

Required: False
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -UnorderedListMarker
Unordered list marker.

```yaml
Type: String
Parameter Sets: Document
Aliases: None
Possible values: -, *, +

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -WriteOptions
Native save settings, or a hashtable of native option properties.

```yaml
Type: MarkdownSaveOptions
Parameter Sets: Document
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
Parameter Sets: Document
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

- `OfficeIMO.Markdown.MarkdownDocument`

## OUTPUTS

- `System.IO.FileInfo`

## RELATED LINKS

- None
