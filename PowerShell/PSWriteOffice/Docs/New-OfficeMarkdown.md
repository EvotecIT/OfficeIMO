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
### __AllParameterSets
```powershell
New-OfficeMarkdown [[-Path] <string>] [[-Content] <scriptblock>] [-Force] [-WriteOptions <MarkdownSaveOptions>] [-WriteProfile <OfficeMarkdownWriteProfile>] [-ImageRenderingMode <MarkdownImageRenderingMode>] [-LineEnding <OfficeLineEnding>] [-UnorderedListMarker <string>] [-PassThru] [-AllowLoss] [-Show] [-WhatIf] [-Confirm] [<CommonParameters>]
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
Gets or sets whether native-save loss is explicitly accepted.

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

### -Content
Gets or sets a content block; omission returns the live document.

```yaml
Type: ScriptBlock
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: 1
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Force
Allow an existing destination; a failed block preserves its original contents.

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

### -ImageRenderingMode
Image serialization used when publishing a block.

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
Line endings used when publishing a block.

```yaml
Type: OfficeLineEnding
Parameter Sets: __AllParameterSets
Aliases: None
Possible values: LF, CRLF, CR

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PassThru
Gets or sets whether a completed content block returns the saved FileInfo.

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
Optional associated path; a content block requires a destination.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Show
Gets or sets whether the saved file opens in the system default application.

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

### -UnorderedListMarker
Unordered list marker used when publishing a block.

```yaml
Type: String
Parameter Sets: __AllParameterSets
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
Writer profile used when publishing a block.

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

- `None`

## OUTPUTS

- `System.IO.FileInfo`
- `OfficeIMO.Markdown.MarkdownDocument`

## RELATED LINKS

- None
