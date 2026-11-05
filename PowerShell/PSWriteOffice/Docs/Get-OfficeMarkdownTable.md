---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Get-OfficeMarkdownTable
## SYNOPSIS
Gets Markdown tables from a Markdown document.

## SYNTAX
### Document (Default)
```powershell
Get-OfficeMarkdownTable [-Document <MarkdownDocument>] [-AsObject] [<CommonParameters>]
```

### Path
```powershell
Get-OfficeMarkdownTable [-Path] <string> [-Options <Object[]>] [-AsObject] [<CommonParameters>]
```

### LiteralPath
```powershell
Get-OfficeMarkdownTable -LiteralPath <string> [-Options <Object[]>] [-AsObject] [<CommonParameters>]
```

### Text
```powershell
Get-OfficeMarkdownTable -Text <string> [-Options <Object[]>] [-AsObject] [<CommonParameters>]
```

## DESCRIPTION
Gets Markdown tables from a Markdown document.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> Open-OfficeMarkdown -Path .\Report.md | Get-OfficeMarkdownTable -AsObject
```

Returns table rows as PowerShell objects using the Markdown header row as property names.

## PARAMETERS

### -AsObject
Emit table rows as PowerShell objects instead of raw table blocks.

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

- `OfficeIMO.Markdown.TableBlock`
- `System.Management.Automation.PSObject`

## RELATED LINKS

- None
