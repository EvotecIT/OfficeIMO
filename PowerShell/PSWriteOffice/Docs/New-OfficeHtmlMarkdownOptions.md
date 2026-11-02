---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# New-OfficeHtmlMarkdownOptions
## SYNOPSIS
Creates native HTML-to-Markdown import and writer settings without converting or writing files.

## SYNTAX
### Default (Default)
```powershell
New-OfficeHtmlMarkdownOptions [-BaseOptions <MarkdownHtmlImportOptions>] [-Portable] [-BaseUri <string>] [-IncludeDocumentChrome] [-PreserveScriptsAndStyles] [-DropUnsupportedBlocks] [-DropUnsupportedInlineHtml] [-MaxInputCharacters <Int32>] [-Base64ImageHandling <MarkdownHtmlBase64ImageHandling>] [-Base64ImageOutputDirectory <string>] [-ListingCardMetadataMode <MarkdownHtmlListingCardMetadataMode>] [-MaxTableExpandedColumns <Int32>] [-WriteOptions <MarkdownSaveOptions>] [-WriteProfile <OfficeMarkdownWriteProfile>] [-ImageRenderingMode <MarkdownImageRenderingMode>] [-LineEnding <OfficeLineEnding>] [-UnorderedListMarker <string>] [<CommonParameters>]
```

## DESCRIPTION
Creates native HTML-to-Markdown import and writer settings without converting or writing files.

## EXAMPLES

### EXAMPLE 1
```powershell
$options = New-OfficeHtmlMarkdownOptions -Portable -LineEnding LF
Convert-OfficeDocument -Path ./Report.html -To Markdown -OutputPath ./Report.md -Options $options
```


## PARAMETERS

### -Base64ImageHandling
How data URI images are represented in Markdown.

```yaml
Type: MarkdownHtmlBase64ImageHandling
Parameter Sets: Default
Aliases: None
Possible values: Include, Skip, SaveToFile

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Base64ImageOutputDirectory
Literal image extraction directory; creating options does not create it.

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

### -BaseOptions
Native settings to copy before applying explicit changes.

```yaml
Type: MarkdownHtmlImportOptions
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
Absolute base URI for relative links and images.

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

### -DropUnsupportedBlocks
Convert unsupported block elements to their child content instead of retaining raw HTML.

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

### -DropUnsupportedInlineHtml
Convert unsupported inline elements instead of retaining raw HTML.

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

### -ImageRenderingMode
How the Markdown writer serializes images.

```yaml
Type: MarkdownImageRenderingMode
Parameter Sets: Default
Aliases: None
Possible values: RichMarkdown, PortableMarkdown, Html

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -IncludeDocumentChrome
Include the complete HTML document instead of only its body.

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

### -LineEnding
Output line endings.

```yaml
Type: OfficeLineEnding
Parameter Sets: Default
Aliases: None
Possible values: LF, CRLF, CR

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ListingCardMetadataMode
Preserve or suppress repeated listing-card metadata.

```yaml
Type: MarkdownHtmlListingCardMetadataMode
Parameter Sets: Default
Aliases: None
Possible values: Preserve, SuppressInRepeatedCards

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -MaxInputCharacters
Maximum accepted HTML input characters; null removes a copied limit.

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

### -MaxTableExpandedColumns
Maximum columns after expanding HTML table spans.

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

### -Portable
Start with the native portable profile instead of the OfficeIMO profile.

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

### -PreserveScriptsAndStyles
Retain script and style elements in the generated Markdown.

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

### -UnorderedListMarker
Unordered list marker.

```yaml
Type: String
Parameter Sets: Default
Aliases: None
Possible values: -, *, +

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -WriteOptions
Native Markdown writer settings to copy.

```yaml
Type: MarkdownSaveOptions
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -WriteProfile
Start with a named Markdown writer profile.

```yaml
Type: OfficeMarkdownWriteProfile
Parameter Sets: Default
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

- `OfficeIMO.Markdown.Html.MarkdownHtmlImportOptions`

## RELATED LINKS

- None
