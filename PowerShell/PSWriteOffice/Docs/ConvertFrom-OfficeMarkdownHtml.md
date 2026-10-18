---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# ConvertFrom-OfficeMarkdownHtml
## SYNOPSIS
Converts HTML content to Markdown.

## SYNTAX
### Html (Default)
```powershell
ConvertFrom-OfficeMarkdownHtml [-Html] <string> [-OutputPath <string>] [-AsDocument] [-PassThru] [-Options <HtmlToMarkdownOptions>] [-Portable] [-BaseUri <string>] [-IncludeDocumentChrome] [-PreserveScriptsAndStyles] [-DropUnsupportedBlocks] [-DropUnsupportedInlineHtml] [-MaxInputCharacters <Int32>] [-Base64ImageHandling <HtmlBase64ImageHandling>] [-Base64ImageOutputDirectory <string>] [-ListingCardMetadataMode <HtmlListingCardMetadataMode>] [-MaxTableExpandedColumns <Int32>] [-WriteOptions <MarkdownWriteOptions>] [-WriteProfile <OfficeMarkdownWriteProfile>] [-ImageRenderingMode <MarkdownImageRenderingMode>] [-LineEnding <string>] [-UnorderedListMarker <string>] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### Path
```powershell
ConvertFrom-OfficeMarkdownHtml [-Path] <string> [-OutputPath <string>] [-AsDocument] [-PassThru] [-Options <HtmlToMarkdownOptions>] [-Portable] [-BaseUri <string>] [-IncludeDocumentChrome] [-PreserveScriptsAndStyles] [-DropUnsupportedBlocks] [-DropUnsupportedInlineHtml] [-MaxInputCharacters <Int32>] [-Base64ImageHandling <HtmlBase64ImageHandling>] [-Base64ImageOutputDirectory <string>] [-ListingCardMetadataMode <HtmlListingCardMetadataMode>] [-MaxTableExpandedColumns <Int32>] [-WriteOptions <MarkdownWriteOptions>] [-WriteProfile <OfficeMarkdownWriteProfile>] [-ImageRenderingMode <MarkdownImageRenderingMode>] [-LineEnding <string>] [-UnorderedListMarker <string>] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Returns Markdown text or saves it to a file when -OutputPath is specified.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $markdown = ConvertFrom-OfficeMarkdownHtml -Html '<h1>Report</h1><p>Ready</p>'
```

Returns Markdown text converted from the supplied HTML.

### EXAMPLE 2
```powershell
PS> $doc = ConvertFrom-OfficeMarkdownHtml -Path .\report.html -AsDocument
```

Returns a Markdown document for further editing or rendering.

## PARAMETERS

### -AsDocument
Emit a Markdown document object instead of Markdown text.

```yaml
Type: SwitchParameter
Parameter Sets: Html, Path
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Base64ImageHandling
Controls how base64 data URI images are converted.

```yaml
Type: HtmlBase64ImageHandling
Parameter Sets: Html, Path
Aliases: None
Possible values: Include, Skip, SaveToFile

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Base64ImageOutputDirectory
Output directory for decoded base64 images when saving them to files.

```yaml
Type: String
Parameter Sets: Html, Path
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -BaseUri
Base URI used to resolve relative links and image sources.

```yaml
Type: String
Parameter Sets: Html, Path
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -DropUnsupportedBlocks
Drop unsupported block HTML instead of preserving it as raw HTML.

```yaml
Type: SwitchParameter
Parameter Sets: Html, Path
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -DropUnsupportedInlineHtml
Drop unsupported inline HTML instead of preserving it as raw HTML.

```yaml
Type: SwitchParameter
Parameter Sets: Html, Path
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Html
HTML markup to convert.

```yaml
Type: String
Parameter Sets: Html
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -ImageRenderingMode
Controls how generated Markdown images are serialized.

```yaml
Type: MarkdownImageRenderingMode
Parameter Sets: Html, Path
Aliases: None
Possible values: RichMarkdown, PortableMarkdown, Html

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -IncludeDocumentChrome
Convert the full HTML document instead of only body contents.

```yaml
Type: SwitchParameter
Parameter Sets: Html, Path
Aliases: None
Possible values:

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
Parameter Sets: Html, Path
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ListingCardMetadataMode
Controls whether repeated listing-card metadata is preserved or suppressed.

```yaml
Type: HtmlListingCardMetadataMode
Parameter Sets: Html, Path
Aliases: None
Possible values: Preserve, SuppressInRepeatedCards

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -MaxInputCharacters
Maximum input length, in characters, accepted by the converter.

```yaml
Type: Int32
Parameter Sets: Html, Path
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -MaxTableExpandedColumns
Maximum logical columns produced by expanding HTML table spans.

```yaml
Type: Int32
Parameter Sets: Html, Path
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Options
Optional conversion options.

```yaml
Type: HtmlToMarkdownOptions
Parameter Sets: Html, Path
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -OutputPath
Optional output path for the Markdown file.

```yaml
Type: String
Parameter Sets: Html, Path
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
Parameter Sets: Html, Path
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Path
Path to an HTML file.

```yaml
Type: String
Parameter Sets: Path
Aliases: InputPath, FilePath
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Portable
Use portable Markdown output when Options is not supplied.

```yaml
Type: SwitchParameter
Parameter Sets: Html, Path
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PreserveScriptsAndStyles
Preserve script, style, noscript, and template elements.

```yaml
Type: SwitchParameter
Parameter Sets: Html, Path
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -UnorderedListMarker
Unordered list marker: '-', '*', or '+'.

```yaml
Type: String
Parameter Sets: Html, Path
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -WriteOptions
Optional Markdown writer options for generated Markdown text.

```yaml
Type: MarkdownWriteOptions
Parameter Sets: Html, Path
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -WriteProfile
Friendly Markdown writer profile for generated Markdown text.

```yaml
Type: OfficeMarkdownWriteProfile
Parameter Sets: Html, Path
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

- `System.String`

## OUTPUTS

- `System.String`
- `System.IO.FileInfo`
- `OfficeIMO.Markdown.MarkdownDoc`

## RELATED LINKS

- None
