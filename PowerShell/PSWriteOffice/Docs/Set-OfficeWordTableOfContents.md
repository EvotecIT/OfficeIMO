---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Set-OfficeWordTableOfContents
## SYNOPSIS
Sets properties on a table of contents in a Word document.

## SYNTAX
### TableOfContents
```powershell
Set-OfficeWordTableOfContents [-TableOfContents <WordTableOfContent>] [-Text <string>] [-TextNoContent <string>] [-PassThru] [<CommonParameters>]
```

### Document
```powershell
Set-OfficeWordTableOfContents [-Document <WordDocument>] [-Text <string>] [-TextNoContent <string>] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Sets properties on a table of contents in a Word document.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> New-OfficeWord -Path .\Report.docx {
    Add-OfficeWordTableOfContents
    Set-OfficeWordTableOfContents -Text 'Contents' -TextNoContent 'No entries yet'
    Add-OfficeWordParagraph -Text 'Executive summary' -Style Heading1
    Update-OfficeWordTableOfContents
}
```

Updates TOC display text, adds heading content, and marks the TOC for refresh.

### EXAMPLE 2
```powershell
PS> $doc = Get-OfficeWord -Path .\Report.docx
$doc |
    Get-OfficeWordTableOfContents |
    Set-OfficeWordTableOfContents -Text 'Report contents'
$doc | Save-OfficeWord -Path .\Report-Toc.docx
```

Pipes the OfficeIMO TOC object into the thin setter and saves the document.

## PARAMETERS

### -Document
Document to update when provided explicitly.

```yaml
Type: WordDocument
Parameter Sets: Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -PassThru
Emit the updated table of contents.

```yaml
Type: SwitchParameter
Parameter Sets: TableOfContents, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -TableOfContents
Table of contents to update.

```yaml
Type: WordTableOfContent
Parameter Sets: TableOfContents
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -Text
Heading text for the table of contents.

```yaml
Type: String
Parameter Sets: TableOfContents, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -TextNoContent
Text shown when the table of contents has no entries.

```yaml
Type: String
Parameter Sets: TableOfContents, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `OfficeIMO.Word.WordTableOfContent`
- `OfficeIMO.Word.WordDocument`

## OUTPUTS

- `OfficeIMO.Word.WordTableOfContent`

## RELATED LINKS

- None
