---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# ConvertFrom-OfficeWordMarkdown
## SYNOPSIS
Creates a Word document from Markdown.

## SYNTAX
### Markdown (Default)
```powershell
ConvertFrom-OfficeWordMarkdown [-Markdown] <string> [-OutputPath <string>] [-TemplatePath <string>] [-BookmarkName <string>] [-ContentControlTag <string>] [-ContentControlAlias <string>] [-KeepPlaceholder] [-RenderFrontMatter] [-FontFamily <string>] [-BaseUri <string>] [-AllowLocalImages] [-AllowedImageDirectory <string[]>] [-AllowRemoteImages] [-ReaderOptions <MarkdownReaderOptions>] [-Profile <MarkdownDialectProfile>] [-NormalizeInput <MarkdownInputNormalizationPreset>] [-Theme <OfficeVisualThemeKind>] [-AllowDataUriImages <Boolean>] [-MaxDataUriImageBytes <Int64>] [-PreferNarrativeSingleLineDefinitions] [-FitImagesToPageContentWidth] [-FitImagesToContextWidth] [-MaxImageWidthPixels <Double>] [-MaxImageHeightPixels <Double>] [-MaxImageWidthPercentOfContent <Double>] [-Open] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### Path
```powershell
ConvertFrom-OfficeWordMarkdown [-Path] <string> [-OutputPath <string>] [-TemplatePath <string>] [-BookmarkName <string>] [-ContentControlTag <string>] [-ContentControlAlias <string>] [-KeepPlaceholder] [-RenderFrontMatter] [-FontFamily <string>] [-BaseUri <string>] [-AllowLocalImages] [-AllowedImageDirectory <string[]>] [-AllowRemoteImages] [-ReaderOptions <MarkdownReaderOptions>] [-Profile <MarkdownDialectProfile>] [-NormalizeInput <MarkdownInputNormalizationPreset>] [-Theme <OfficeVisualThemeKind>] [-AllowDataUriImages <Boolean>] [-MaxDataUriImageBytes <Int64>] [-PreferNarrativeSingleLineDefinitions] [-FitImagesToPageContentWidth] [-FitImagesToContextWidth] [-MaxImageWidthPixels <Double>] [-MaxImageHeightPixels <Double>] [-MaxImageWidthPercentOfContent <Double>] [-Open] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### Document
```powershell
ConvertFrom-OfficeWordMarkdown -Document <MarkdownDoc> [-OutputPath <string>] [-TemplatePath <string>] [-BookmarkName <string>] [-ContentControlTag <string>] [-ContentControlAlias <string>] [-KeepPlaceholder] [-RenderFrontMatter] [-FontFamily <string>] [-BaseUri <string>] [-AllowLocalImages] [-AllowedImageDirectory <string[]>] [-AllowRemoteImages] [-ReaderOptions <MarkdownReaderOptions>] [-Profile <MarkdownDialectProfile>] [-NormalizeInput <MarkdownInputNormalizationPreset>] [-Theme <OfficeVisualThemeKind>] [-AllowDataUriImages <Boolean>] [-MaxDataUriImageBytes <Int64>] [-PreferNarrativeSingleLineDefinitions] [-FitImagesToPageContentWidth] [-FitImagesToContextWidth] [-MaxImageWidthPixels <Double>] [-MaxImageHeightPixels <Double>] [-MaxImageWidthPercentOfContent <Double>] [-Open] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Returns a WordDocument or saves it to disk when -OutputPath is provided.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> ConvertFrom-OfficeWordMarkdown -Markdown '# Hello' -OutputPath .\hello.docx
```

Writes a Word document containing the supplied Markdown.

### EXAMPLE 2
```powershell
PS> Get-OfficeMarkdown -Path .\README.md | ConvertFrom-OfficeWordMarkdown
```

Returns a Word document instance for further edits.

### EXAMPLE 3
```powershell
PS> ConvertFrom-OfficeWordMarkdown -Path .\SOP.md -TemplatePath .\Template.docx -BookmarkName MainContent -OutputPath .\SOP.docx
```

Copies the template and replaces the bookmark paragraph with generated Markdown content.

## PARAMETERS

### -AllowDataUriImages
Allow data URI Markdown images to be embedded in Word output.

```yaml
Type: Boolean
Parameter Sets: Markdown, Path, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -AllowedImageDirectory
Restrict local images to one or more directories.

```yaml
Type: String[]
Parameter Sets: Markdown, Path, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -AllowLocalImages
Allow local Markdown images to be inserted into the document.

```yaml
Type: SwitchParameter
Parameter Sets: Markdown, Path, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -AllowRemoteImages
Allow remote HTTP(S) images to be downloaded and inserted.

```yaml
Type: SwitchParameter
Parameter Sets: Markdown, Path, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -BaseUri
Base URI used to resolve relative links and images.

```yaml
Type: String
Parameter Sets: Markdown, Path, Document
Aliases: BasePath
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -BookmarkName
Bookmark name that marks where Markdown content should be inserted in the template.

```yaml
Type: String
Parameter Sets: Markdown, Path, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ContentControlAlias
Block content control alias that marks where Markdown content should be inserted in the template.

```yaml
Type: String
Parameter Sets: Markdown, Path, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ContentControlTag
Block content control tag that marks where Markdown content should be inserted in the template.

```yaml
Type: String
Parameter Sets: Markdown, Path, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Document
Markdown document instance to convert.

```yaml
Type: MarkdownDoc
Parameter Sets: Document
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -FitImagesToContextWidth
Fit Markdown images to the current content context width.

```yaml
Type: SwitchParameter
Parameter Sets: Markdown, Path, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -FitImagesToPageContentWidth
Fit Markdown images to the page content width.

```yaml
Type: SwitchParameter
Parameter Sets: Markdown, Path, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -FontFamily
Optional font family applied during conversion.

```yaml
Type: String
Parameter Sets: Markdown, Path, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -KeepPlaceholder
Keep the target bookmark or content-control placeholder after inserting Markdown content.

```yaml
Type: SwitchParameter
Parameter Sets: Markdown, Path, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Markdown
Markdown text to convert.

```yaml
Type: String
Parameter Sets: Markdown
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -MaxDataUriImageBytes
Maximum decoded size for one data URI image.

```yaml
Type: Int64
Parameter Sets: Markdown, Path, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -MaxImageHeightPixels
Optional hard cap for Markdown image height in pixels.

```yaml
Type: Double
Parameter Sets: Markdown, Path, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -MaxImageWidthPercentOfContent
Optional hard cap for Markdown image width as a percentage of available content width.

```yaml
Type: Double
Parameter Sets: Markdown, Path, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -MaxImageWidthPixels
Optional hard cap for Markdown image width in pixels.

```yaml
Type: Double
Parameter Sets: Markdown, Path, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NormalizeInput
Applies a built-in Markdown input normalization preset before parsing.

```yaml
Type: MarkdownInputNormalizationPreset
Parameter Sets: Markdown, Path, Document
Aliases: None
Possible values: None, IntelligenceXTranscript, IntelligenceXTranscriptStrict, DocsLoose

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Open
Open the document after saving.

```yaml
Type: SwitchParameter
Parameter Sets: Markdown, Path, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -OutputPath
Optional output path for the .docx file.

```yaml
Type: String
Parameter Sets: Markdown, Path, Document
Aliases: OutPath
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PassThru
Emit a FileInfo when saving to disk.

```yaml
Type: SwitchParameter
Parameter Sets: Markdown, Path, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Path
Path to a Markdown file.

```yaml
Type: String
Parameter Sets: Path
Aliases: FilePath
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PreferNarrativeSingleLineDefinitions
Prefer narrative paragraphs for isolated single-line definition-list patterns.

```yaml
Type: SwitchParameter
Parameter Sets: Markdown, Path, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Profile
Named Markdown reader profile used when ReaderOptions is not supplied.

```yaml
Type: MarkdownDialectProfile
Parameter Sets: Markdown, Path, Document
Aliases: None
Possible values: OfficeIMO, CommonMark, GitHubFlavoredMarkdown, Portable

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ReaderOptions
Optional Markdown reader options used before Word conversion.

```yaml
Type: MarkdownReaderOptions
Parameter Sets: Markdown, Path, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -RenderFrontMatter
Render YAML front matter as visible Word content. Template conversions hide front matter by default.

```yaml
Type: SwitchParameter
Parameter Sets: Markdown, Path, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -TemplatePath
Optional Word template document to copy before inserting Markdown content.

```yaml
Type: String
Parameter Sets: Markdown, Path, Document
Aliases: Template
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Theme
Shared Markdown visual theme for generated Word output.

```yaml
Type: OfficeVisualThemeKind
Parameter Sets: Markdown, Path, Document
Aliases: None
Possible values: Plain, WordLike, TechnicalDocument, GitHubLike, Compact, Report

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `System.String`
- `OfficeIMO.Markdown.MarkdownDoc`

## OUTPUTS

- `OfficeIMO.Word.WordDocument`
- `System.IO.FileInfo`

## RELATED LINKS

- None
