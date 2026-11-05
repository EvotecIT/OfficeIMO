---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Get-OfficeMarkdownNode
## SYNOPSIS
Gets the OfficeIMO.Markdown object tree from Markdown content.

## SYNTAX
### Document (Default)
```powershell
Get-OfficeMarkdownNode [-Document <MarkdownDocument>] [-NodeType <string>] [-MaxDepth <int>] [-CaseSensitive] [-Raw] [<CommonParameters>]
```

### Path
```powershell
Get-OfficeMarkdownNode [-Path] <string> [-Options <Object[]>] [-NodeType <string>] [-MaxDepth <int>] [-CaseSensitive] [-Raw] [<CommonParameters>]
```

### LiteralPath
```powershell
Get-OfficeMarkdownNode -LiteralPath <string> [-Options <Object[]>] [-NodeType <string>] [-MaxDepth <int>] [-CaseSensitive] [-Raw] [<CommonParameters>]
```

### Text
```powershell
Get-OfficeMarkdownNode -Text <string> [-Options <Object[]>] [-NodeType <string>] [-MaxDepth <int>] [-CaseSensitive] [-Raw] [<CommonParameters>]
```

## DESCRIPTION
Returns PowerShell-friendly node records by default. Use -Raw to emit the underlying OfficeIMO nodes.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> Get-OfficeMarkdownNode -Text "# Report`n`n## Summary"
```

Parses Markdown text and returns the document, block, and inline object tree.

### EXAMPLE 2
```powershell
PS> Open-OfficeMarkdown -Path .\README.md | Get-OfficeMarkdownNode -NodeType '*Table*'
```

Returns matching nodes from an existing Markdown document.

## PARAMETERS

### -CaseSensitive
Use case-sensitive matching for node type filters.

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

### -MaxDepth
Maximum traversal depth. Zero returns only the document root.

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

### -NodeType
Optional wildcard pattern matched against node type names.

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

### -Raw
Emit raw OfficeIMO.Markdown node objects instead of PowerShell-friendly records.

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
- `OfficeIMO.Markdown.MarkdownObject`

## RELATED LINKS

- None
