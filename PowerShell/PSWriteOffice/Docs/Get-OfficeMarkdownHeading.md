---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Get-OfficeMarkdownHeading
## SYNOPSIS
Gets heading metadata from a Markdown document.

## SYNTAX
### Path (Default)
```powershell
Get-OfficeMarkdownHeading [-Path] <string> [-Options <MarkdownReaderOptions>] [-Profile <MarkdownDialectProfile>] [-BaseUri <string>] [-MaxInputCharacters <Int32>] [-NormalizeInput <MarkdownInputNormalizationPreset>] [-DisallowFileUrls <Boolean>] [-AllowDataUrls <Boolean>] [-AllowMailtoUrls <Boolean>] [-AllowProtocolRelativeUrls <Boolean>] [-RestrictUrlSchemes <Boolean>] [-AllowedUrlScheme <string[]>] [-MinLevel <int>] [-MaxLevel <int>] [-HeadingText <string>] [-Anchor <string>] [-CaseSensitive] [<CommonParameters>]
```

### Document
```powershell
Get-OfficeMarkdownHeading -Document <MarkdownDoc> [-Options <MarkdownReaderOptions>] [-Profile <MarkdownDialectProfile>] [-BaseUri <string>] [-MaxInputCharacters <Int32>] [-NormalizeInput <MarkdownInputNormalizationPreset>] [-DisallowFileUrls <Boolean>] [-AllowDataUrls <Boolean>] [-AllowMailtoUrls <Boolean>] [-AllowProtocolRelativeUrls <Boolean>] [-RestrictUrlSchemes <Boolean>] [-AllowedUrlScheme <string[]>] [-MinLevel <int>] [-MaxLevel <int>] [-HeadingText <string>] [-Anchor <string>] [-CaseSensitive] [<CommonParameters>]
```

### Text
```powershell
Get-OfficeMarkdownHeading -Text <string> [-Options <MarkdownReaderOptions>] [-Profile <MarkdownDialectProfile>] [-BaseUri <string>] [-MaxInputCharacters <Int32>] [-NormalizeInput <MarkdownInputNormalizationPreset>] [-DisallowFileUrls <Boolean>] [-AllowDataUrls <Boolean>] [-AllowMailtoUrls <Boolean>] [-AllowProtocolRelativeUrls <Boolean>] [-RestrictUrlSchemes <Boolean>] [-AllowedUrlScheme <string[]>] [-MinLevel <int>] [-MaxLevel <int>] [-HeadingText <string>] [-Anchor <string>] [-CaseSensitive] [<CommonParameters>]
```

## DESCRIPTION
Returns heading level, text, resolved anchor, and the backing heading block.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> Get-OfficeMarkdownHeading -Text "# Title`n`n## Details"
```

Parses Markdown text and returns the document headings.

### EXAMPLE 2
```powershell
PS> Get-OfficeMarkdown -Path .\README.md | Get-OfficeMarkdownHeading -MinLevel 2
```

Returns headings from an existing Markdown document object.

## PARAMETERS

### -AllowDataUrls
Allow data URLs while parsing Markdown links and images.

```yaml
Type: Boolean
Parameter Sets: Path, Document, Text
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -AllowedUrlScheme
Allowed URL schemes when URL scheme restriction is enabled.

```yaml
Type: String[]
Parameter Sets: Path, Document, Text
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -AllowMailtoUrls
Allow mailto URLs while parsing Markdown links.

```yaml
Type: Boolean
Parameter Sets: Path, Document, Text
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -AllowProtocolRelativeUrls
Allow protocol-relative URLs while parsing Markdown links and images.

```yaml
Type: Boolean
Parameter Sets: Path, Document, Text
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Anchor
Optional wildcard pattern matched against resolved heading anchors.

```yaml
Type: String
Parameter Sets: Path, Document, Text
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -BaseUri
Base URI used to resolve and restrict relative Markdown links and images.

```yaml
Type: String
Parameter Sets: Path, Document, Text
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -CaseSensitive
Use case-sensitive matching for text and anchor filters.

```yaml
Type: SwitchParameter
Parameter Sets: Path, Document, Text
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -DisallowFileUrls
Block file URLs while parsing Markdown links and images.

```yaml
Type: Boolean
Parameter Sets: Path, Document, Text
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Document
Markdown document to inspect.

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

### -HeadingText
Optional wildcard pattern matched against heading text.

```yaml
Type: String
Parameter Sets: Path, Document, Text
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -MaxInputCharacters
Maximum Markdown input length accepted by the reader.

```yaml
Type: Int32
Parameter Sets: Path, Document, Text
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -MaxLevel
Maximum heading level to return.

```yaml
Type: Int32
Parameter Sets: Path, Document, Text
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -MinLevel
Minimum heading level to return.

```yaml
Type: Int32
Parameter Sets: Path, Document, Text
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
Parameter Sets: Path, Document, Text
Aliases: None
Possible values: None, IntelligenceXTranscript, IntelligenceXTranscriptStrict, DocsLoose

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Options
Optional reader options used when parsing path or text input.

```yaml
Type: MarkdownReaderOptions
Parameter Sets: Path, Document, Text
Aliases: ReaderOptions
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Path
Path to the Markdown file.

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

### -Profile
Named reader profile used when Options is not supplied.

```yaml
Type: MarkdownDialectProfile
Parameter Sets: Path, Document, Text
Aliases: None
Possible values: OfficeIMO, CommonMark, GitHubFlavoredMarkdown, Portable

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -RestrictUrlSchemes
Restrict parsed URL schemes to the allow-list.

```yaml
Type: Boolean
Parameter Sets: Path, Document, Text
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Text
Markdown text to parse.

```yaml
Type: String
Parameter Sets: Text
Aliases: None
Possible values:

Required: True
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

- `OfficeIMO.Markdown.MarkdownDoc+HeadingInfo`

## RELATED LINKS

- None
