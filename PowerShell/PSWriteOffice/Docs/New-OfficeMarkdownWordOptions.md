---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# New-OfficeMarkdownWordOptions
## SYNOPSIS
Creates native Markdown-to-Word settings, including an engine-owned template path.

## SYNTAX
### Default (Default)
```powershell
New-OfficeMarkdownWordOptions [-Template <string>] [-BookmarkName <string>] [-ContentControlTag <string>] [-ContentControlAlias <string>] [-KeepPlaceholder] [-NoFrontMatter] [-Font <string>] [-BaseUri <string>] [-AllowLocalImages] [-AllowedImageDirectory <string[]>] [-Theme <OfficeThemeInput>] [-AllowRemoteImages] [-ReaderOptions <MarkdownLoadOptions>] [-Profile <MarkdownDialectProfile>] [-NormalizeInput <MarkdownInputNormalizationPreset>] [-NoDataUriImages] [-MaxDataUriImageBytes <Int64>] [-PreferNarrativeSingleLineDefinitions] [-FitImagesToPageContentWidth] [-FitImagesToContextWidth] [-MaxImageWidth <OfficeLength>] [-MaxImageHeight <OfficeLength>] [-MaxImageWidthPercentOfContent <Double>] [<CommonParameters>]
```

## DESCRIPTION
Creates native Markdown-to-Word settings, including an engine-owned template path.

## EXAMPLES

### EXAMPLE 1
```powershell
$options = New-OfficeMarkdownWordOptions -Template .\Templates\Runbook.dotx -BookmarkName Body
Convert-OfficeDocument -Path .\Runbook.md -To Word -OutputPath .\Runbook.docx -Options $options
```


## PARAMETERS

### -AllowedImageDirectory
AllowedImageDirectory override; omission preserves the native default.

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
AllowLocalImages override; omission preserves the native default.

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
Allow the native bounded HTTP image loader. Omission keeps remote images disabled.

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

### -BaseUri
BaseUri override; omission preserves the native default.

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

### -BookmarkName
BookmarkName override; omission preserves the native default.

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

### -ContentControlAlias
ContentControlAlias override; omission preserves the native default.

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

### -ContentControlTag
ContentControlTag override; omission preserves the native default.

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

### -FitImagesToContextWidth
FitImagesToContextWidth override; omission preserves the native default.

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

### -FitImagesToPageContentWidth
FitImagesToPageContentWidth override; omission preserves the native default.

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
Font override; omission preserves the native default.

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

### -KeepPlaceholder
KeepPlaceholder override; omission preserves the native default.

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

### -MaxDataUriImageBytes
MaxDataUriImageBytes override; omission preserves the native default.

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

### -MaxImageHeight
MaxImageHeight override; omission preserves the native default.

```yaml
Type: OfficeLength
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -MaxImageWidth
MaxImageWidth override; omission preserves the native default.

```yaml
Type: OfficeLength
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -MaxImageWidthPercentOfContent
MaxImageWidthPercentOfContent override; omission preserves the native default.

```yaml
Type: Double
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
NoDataUriImages override; omission preserves the native default.

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
NoFrontMatter override; omission preserves the native default.

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

### -NormalizeInput
NormalizeInput override; omission preserves the native default.

```yaml
Type: MarkdownInputNormalizationPreset
Parameter Sets: Default
Aliases: None
Possible values: None, IntelligenceXTranscript, IntelligenceXTranscriptStrict, DocsLoose

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PreferNarrativeSingleLineDefinitions
PreferNarrativeSingleLineDefinitions override; omission preserves the native default.

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
Profile override; omission preserves the native default.

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
ReaderOptions override; omission preserves the native default.

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

### -Template
Template override; omission preserves the native default.

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

### -Theme
Built-in theme, shared theme object or theme template.

```yaml
Type: OfficeThemeInput
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

- `OfficeIMO.Word.Markdown.MarkdownWordExportOptions`

## RELATED LINKS

- None
