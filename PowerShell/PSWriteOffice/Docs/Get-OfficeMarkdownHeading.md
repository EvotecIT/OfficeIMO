---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Get-OfficeMarkdownHeading
## SYNOPSIS
Gets heading metadata from a Markdown document.

## SYNTAX
### Document (Default)
```powershell
Get-OfficeMarkdownHeading [-Document <MarkdownDocument>] [-MinLevel <int>] [-MaxLevel <int>] [-HeadingText <string>] [-Anchor <string>] [-CaseSensitive] [<CommonParameters>]
```

### Path
```powershell
Get-OfficeMarkdownHeading [-Path] <string> [-Options <Object[]>] [-MinLevel <int>] [-MaxLevel <int>] [-HeadingText <string>] [-Anchor <string>] [-CaseSensitive] [<CommonParameters>]
```

### LiteralPath
```powershell
Get-OfficeMarkdownHeading -LiteralPath <string> [-Options <Object[]>] [-MinLevel <int>] [-MaxLevel <int>] [-HeadingText <string>] [-Anchor <string>] [-CaseSensitive] [<CommonParameters>]
```

### Text
```powershell
Get-OfficeMarkdownHeading -Text <string> [-Options <Object[]>] [-MinLevel <int>] [-MaxLevel <int>] [-HeadingText <string>] [-Anchor <string>] [-CaseSensitive] [<CommonParameters>]
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
PS> Open-OfficeMarkdown -Path .\README.md | Get-OfficeMarkdownHeading -MinLevel 2
```

Returns headings from an existing Markdown document object.

## PARAMETERS

### -Anchor
Optional wildcard pattern matched against resolved heading anchors.

```yaml
Type: String
Parameter Sets: Document, Path, LiteralPath, Text
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
Parameter Sets: Document, Path, LiteralPath, Text
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Document
Caller-owned document; omission uses the enclosing Markdown document.

```yaml
Type: MarkdownDocument
Parameter Sets: Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -HeadingText
Optional wildcard pattern matched against heading text.

```yaml
Type: String
Parameter Sets: Document, Path, LiteralPath, Text
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -LiteralPath
Literal input file, including a FileInfo pipeline.

```yaml
Type: String
Parameter Sets: LiteralPath
Aliases: FullName, PSPath
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: True (ByPropertyName)
Accept wildcard characters: False
```

### -MaxLevel
Maximum heading level to return.

```yaml
Type: Int32
Parameter Sets: Document, Path, LiteralPath, Text
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
Parameter Sets: Document, Path, LiteralPath, Text
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Options
Native reader options and optional native-property hashtable overrides.

```yaml
Type: Object[]
Parameter Sets: Path, LiteralPath, Text
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Path
Input file with FileSystem wildcard expansion.

```yaml
Type: String
Parameter Sets: Path
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: True
```

### -Text
Markdown source text.

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

- `OfficeIMO.Markdown.MarkdownDocument`
- `System.String`

## OUTPUTS

- `OfficeIMO.Markdown.MarkdownDocument+HeadingInfo`

## RELATED LINKS

- None
