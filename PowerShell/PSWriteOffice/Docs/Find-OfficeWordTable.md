---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Find-OfficeWordTable
## SYNOPSIS
Finds Word tables containing matching cell text.

## SYNTAX
### PathText (Default)
```powershell
Find-OfficeWordTable [-Path] <string> [-Text] <string> [-CaseSensitive] [-IncludeNested] [<CommonParameters>]
```

### PathRegex
```powershell
Find-OfficeWordTable [-Path] <string> [-Pattern] <string> [-CaseSensitive] [-IncludeNested] [<CommonParameters>]
```

### DocumentText
```powershell
Find-OfficeWordTable [-Text] <string> -Document <WordDocument> [-CaseSensitive] [-IncludeNested] [<CommonParameters>]
```

### DocumentRegex
```powershell
Find-OfficeWordTable [-Pattern] <string> -Document <WordDocument> [-CaseSensitive] [-IncludeNested] [<CommonParameters>]
```

## DESCRIPTION
Searches table cell paragraphs in a Word document and returns the matching WordTable
objects. Use this when a document came from a template or another system and the script needs to
locate a table by visible marker text before appending rows, changing cells, or applying table-cell
formatting.

By default only top-level tables are searched. Use -IncludeNested to include tables inside
table cells. Use -Text for literal contains matching or -Pattern for regular
expressions.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $doc = Get-OfficeWord -Path .\Report.docx
$table = Find-OfficeWordTable -Document $doc -Text 'Risk register' | Select-Object -First 1
$table | Add-OfficeWordTableRow -Values 'Contoso', 'Open', 'High'
```

Searches table cell paragraphs and returns the matching OfficeIMO table object.

### EXAMPLE 2
```powershell
PS> $doc = Get-OfficeWord -Path .\Handover.docx
$table = Find-OfficeWordTable -Document $doc -Text 'Risk marker' | Select-Object -First 1
$table | Add-OfficeWordTableRow -Values 'Mitigation plan', 'Service Desk', 'Ready'
$table | Get-OfficeWordTableCell -Row 2 -Column 2 |
    Set-OfficeWordTableCell -Text 'Investigating' -ShadingFillColor '#fff2cc' -ShadingPattern Clear
$doc | Close-OfficeWord -Save
```

Shows the common existing-document workflow: locate the table, mutate it, then save the open document.

## PARAMETERS

### -CaseSensitive
Use case-sensitive matching.

```yaml
Type: SwitchParameter
Parameter Sets: PathText, PathRegex, DocumentText, DocumentRegex
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Document
Open document to inspect. The caller controls the document lifetime.

```yaml
Type: WordDocument
Parameter Sets: DocumentText, DocumentRegex
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -IncludeNested
Include nested tables inside table cells.

```yaml
Type: SwitchParameter
Parameter Sets: PathText, PathRegex, DocumentText, DocumentRegex
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Path
Path to the document to open read-only for searching.

```yaml
Type: String
Parameter Sets: PathText, PathRegex
Aliases: InputPath, FilePath
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Pattern
Regular expression pattern to find in table cells.

```yaml
Type: String
Parameter Sets: PathRegex, DocumentRegex
Aliases: None
Possible values:

Required: True
Position: 1
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Text
Literal text to find in table cells.

```yaml
Type: String
Parameter Sets: PathText, DocumentText
Aliases: None
Possible values:

Required: True
Position: 1
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `OfficeIMO.Word.WordDocument`

## OUTPUTS

- `OfficeIMO.Word.WordTable`

## RELATED LINKS

- None
