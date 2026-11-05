---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# New-OfficeMarkdownPowerPointOptions
## SYNOPSIS
Creates native Markdown slide boundaries, speaker-note, front-matter and image policies.

## SYNTAX
### Default (Default)
```powershell
New-OfficeMarkdownPowerPointOptions [-SlideBreak <MarkdownSlideBreak>] [-TitleHeadingLevel <Int32>] [-NoBlockquoteNotes] [-NoFrontMatter] [-SlideSize <PowerPointSlideSizePreset>] [-BaseDirectory <string>] [-AllowLocalImages] [-AllowedImageDirectory <string[]>] [-AllowRemoteImages] [-NoDataUriImages] [-MaximumDataUriImageBytes <Int64>] [-Profile <MarkdownDialectProfile>] [-ReaderOptions <MarkdownLoadOptions>] [<CommonParameters>]
```

## DESCRIPTION
Creates native Markdown slide boundaries, speaker-note, front-matter and image policies.

## EXAMPLES

### EXAMPLE 1
```powershell
$options = New-OfficeMarkdownPowerPointOptions -SlideBreak Heading -TitleHeadingLevel 2
Convert-OfficeDocument -Path ./Kickoff.md -To PowerPoint -OutputPath ./Kickoff.pptx -Options $options
```


## PARAMETERS

### -AllowedImageDirectory
Allowed local image directories.

```yaml
Type: String[]
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -AllowLocalImages
Allow native local image loading within the configured roots.

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

### -AllowRemoteImages
Allow the native bounded HTTP image loader.

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

### -BaseDirectory
Base directory for relative images.

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

### -MaximumDataUriImageBytes
Maximum decoded data URI image bytes.

```yaml
Type: Int64
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NoBlockquoteNotes
Keep blockquotes visible instead of placing them in speaker notes.

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

### -NoDataUriImages
Disable data URI images.

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

### -NoFrontMatter
Ignore title, subtitle and author front matter.

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

### -Profile
Native Markdown dialect profile; mutually exclusive with ReaderOptions.

```yaml
Type: MarkdownDialectProfile
Parameter Sets: Default
Aliases: None
Possible values: OfficeIMO, CommonMark, GitHubFlavoredMarkdown, Portable

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ReaderOptions
Native parser, normalization and input-limit policy, captured independently.

```yaml
Type: MarkdownLoadOptions
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -SlideBreak
Native boundaries that begin a slide.

```yaml
Type: MarkdownSlideBreak
Parameter Sets: Default
Aliases: None
Possible values: ThematicBreakOrHeading, ThematicBreak, Heading

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -SlideSize
Native presentation canvas preset.

```yaml
Type: PowerPointSlideSizePreset
Parameter Sets: Default
Aliases: None
Possible values: Screen4x3, Screen16x9, Screen16x10

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -TitleHeadingLevel
Heading level used for slide titles.

```yaml
Type: Int32
Parameter Sets: Default
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

- `None`

## OUTPUTS

- `OfficeIMO.Workflows.MarkdownToPowerPointOptions`

## RELATED LINKS

- None
