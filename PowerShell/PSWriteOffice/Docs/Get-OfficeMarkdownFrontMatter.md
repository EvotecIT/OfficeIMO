---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Get-OfficeMarkdownFrontMatter
## SYNOPSIS
Gets top-level YAML or TOML front matter entries from a Markdown document as Key, Value, Node and SourceSpan objects.

## SYNTAX
### Document (Default)
```powershell
Get-OfficeMarkdownFrontMatter [-Document <MarkdownDocument>] [-Key <string>] [-CaseSensitive] [<CommonParameters>]
```

### Path
```powershell
Get-OfficeMarkdownFrontMatter [-Path] <string> [-Options <Object[]>] [-Key <string>] [-CaseSensitive] [<CommonParameters>]
```

### LiteralPath
```powershell
Get-OfficeMarkdownFrontMatter -LiteralPath <string> [-Options <Object[]>] [-Key <string>] [-CaseSensitive] [<CommonParameters>]
```

### Text
```powershell
Get-OfficeMarkdownFrontMatter -Text <string> [-Options <Object[]>] [-Key <string>] [-CaseSensitive] [<CommonParameters>]
```

## DESCRIPTION
Gets YAML front matter entries from a Markdown document.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $metadata = Get-OfficeMarkdownFrontMatter -Path .\Report.md -Key 'title'
$metadata |
    Select-Object -Property Key, Value |
    Format-Table -AutoSize
```

Parses a Markdown file and returns matching front matter entries for metadata proof.

### EXAMPLE 2
```powershell
PS> $markdown = "---`ntitle: Report`nstatus: Ready`n---`n# Report"
Get-OfficeMarkdownFrontMatter -Text $markdown |
    Select-Object -Property Key, Value
```

Parses Markdown text directly when the document has not been saved yet.

## PARAMETERS

### -CaseSensitive
Use case-sensitive matching for key filters.

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

### -Key
Optional wildcard pattern matched against front matter keys.

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

- `System.Management.Automation.PSObject`

## RELATED LINKS

- None
