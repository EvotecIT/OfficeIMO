---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# New-OfficeMarkdown
## SYNOPSIS
Creates a live Markdown document, or atomically publishes a completed content block.

## SYNTAX
### Live (Default)
```powershell
New-OfficeMarkdown [[-Path] <string>] [-Force] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### Content
```powershell
New-OfficeMarkdown [-Path] <string> [-Content] <scriptblock> [-Force] [-WriteOptions <MarkdownSaveOptions>] [-WriteProfile <OfficeMarkdownWriteProfile>] [-ImageRenderingMode <MarkdownImageRenderingMode>] [-LineEnding <OfficeLineEnding>] [-UnorderedListMarker <string>] [-PassThru] [-AllowLoss] [-Show] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Creates a live Markdown document, or atomically publishes a completed content block.

## EXAMPLES

### EXAMPLE 1
```powershell
New-OfficeMarkdown -Path ./Report.md -PassThru { MarkdownHeading -Text 'Report' }
```


### EXAMPLE 2
```powershell
$document = New-OfficeMarkdown -Path ./Report.md
$document | MarkdownHeading -Text 'Report'
$document | Save-OfficeMarkdown
$document | Close-OfficeMarkdown
```


## PARAMETERS

### -AllowLoss
Explicitly accept reductions reported by the content block's native save.

```yaml
Type: SwitchParameter
Parameter Sets: Content
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Content
Authoring block to complete before the document is atomically saved and closed.

```yaml
Type: ScriptBlock
Parameter Sets: Content
Aliases: None
Possible values:

Required: True
Position: 1
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Force
Allow an existing destination; a failed block preserves its original contents.

```yaml
Type: SwitchParameter
Parameter Sets: Live, Content
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ImageRenderingMode
Image serialization used when publishing a block.

```yaml
Type: MarkdownImageRenderingMode
Parameter Sets: Content
Aliases: None
Possible values: RichMarkdown, PortableMarkdown, Html

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -LineEnding
Line endings used when publishing a block.

```yaml
Type: OfficeLineEnding
Parameter Sets: Content
Aliases: None
Possible values: LF, CRLF, CR

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PassThru
Return the saved FileInfo after a completed content block.

```yaml
Type: SwitchParameter
Parameter Sets: Content
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Path
Optional associated destination for a live document; required when a content block saves a file.

```yaml
Type: String
Parameter Sets: Live, Content
Aliases: None
Possible values:

Required: False
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Show
Open the saved result in its registered application after the content block succeeds.

```yaml
Type: SwitchParameter
Parameter Sets: Content
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -UnorderedListMarker
Unordered list marker used when publishing a block.

```yaml
Type: String
Parameter Sets: Content
Aliases: None
Possible values: -, *, +

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -WriteOptions
Native Markdown save settings or a hashtable, used when publishing a block.

```yaml
Type: MarkdownSaveOptions
Parameter Sets: Content
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -WriteProfile
Writer profile used when publishing a block.

```yaml
Type: OfficeMarkdownWriteProfile
Parameter Sets: Content
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

- `None`

## OUTPUTS

- `System.IO.FileInfo`
- `OfficeIMO.Markdown.MarkdownDocument`

## RELATED LINKS

- None
