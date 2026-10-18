---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Get-OfficeWordHyperlink
## SYNOPSIS
Gets hyperlinks from a Word document.

## SYNTAX
### Path (Default)
```powershell
Get-OfficeWordHyperlink [-Path] <string> [-Text <string[]>] [-Url <string[]>] [-Anchor <string[]>] [<CommonParameters>]
```

### Document
```powershell
Get-OfficeWordHyperlink -Document <WordDocument> [-Text <string[]>] [-Url <string[]>] [-Anchor <string[]>] [<CommonParameters>]
```

### Section
```powershell
Get-OfficeWordHyperlink -Section <WordSection> [-Text <string[]>] [-Url <string[]>] [-Anchor <string[]>] [<CommonParameters>]
```

### Paragraph
```powershell
Get-OfficeWordHyperlink -Paragraph <WordParagraph> [-Text <string[]>] [-Url <string[]>] [-Anchor <string[]>] [<CommonParameters>]
```

## DESCRIPTION
Gets hyperlinks from a Word document.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $links = Get-OfficeWordHyperlink -Path .\Report.docx -Url 'https://*'
$links |
    Select-Object -Property Text, Uri |
    Export-Csv -Path .\ReportLinks.csv -NoTypeInformation
```

Filters external hyperlinks and exports them for link validation or review.

## PARAMETERS

### -Anchor
Filter by bookmark anchor (wildcards supported).

```yaml
Type: String[]
Parameter Sets: Path, Document, Section, Paragraph
Aliases: Bookmark
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Document
Document to inspect.

```yaml
Type: WordDocument
Parameter Sets: Document
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -Paragraph
Paragraph to inspect.

```yaml
Type: WordParagraph
Parameter Sets: Paragraph
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -Path
Path to the document.

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

### -Section
Section to inspect.

```yaml
Type: WordSection
Parameter Sets: Section
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -Text
Filter by hyperlink text (wildcards supported).

```yaml
Type: String[]
Parameter Sets: Path, Document, Section, Paragraph
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Url
Filter by external URL (wildcards supported).

```yaml
Type: String[]
Parameter Sets: Path, Document, Section, Paragraph
Aliases: Uri
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `OfficeIMO.Word.WordDocument`
- `OfficeIMO.Word.WordSection`
- `OfficeIMO.Word.WordParagraph`

## OUTPUTS

- `OfficeIMO.Word.WordHyperLink`

## RELATED LINKS

- None
