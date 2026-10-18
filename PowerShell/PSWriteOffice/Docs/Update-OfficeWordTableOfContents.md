---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Update-OfficeWordTableOfContents
## SYNOPSIS
Updates the table of contents in a Word document.

## SYNTAX
### Document (Default)
```powershell
Update-OfficeWordTableOfContents [-Document <WordDocument>] [-Regenerate] [-PassThru] [<CommonParameters>]
```

### TableOfContents
```powershell
Update-OfficeWordTableOfContents [-TableOfContents <WordTableOfContent>] [-Regenerate] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Updates the table of contents in a Word document.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> New-OfficeWord -Path .\ExecutiveReport.docx {
    Add-OfficeWordTableOfContents
    Add-OfficeWordParagraph -Text 'Executive summary' -Style Heading1
    Add-OfficeWordParagraph -Text 'Summary text'
    Update-OfficeWordTableOfContents
}
```

Marks TOC fields as dirty and updates the document settings so Word refreshes the TOC when opened.

### EXAMPLE 2
```powershell
PS> $doc = Get-OfficeWord -Path .\Report.docx
$doc | Update-OfficeWordTableOfContents -Regenerate
$doc | Save-OfficeWord -Path .\Report-RegeneratedToc.docx
```

Uses OfficeIMO's regenerate path, then saves the updated document.

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
Parameter Sets: Document, TableOfContents
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Regenerate
Rebuild the table of contents before updating.

```yaml
Type: SwitchParameter
Parameter Sets: Document, TableOfContents
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

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `OfficeIMO.Word.WordTableOfContent`
- `OfficeIMO.Word.WordDocument`

## OUTPUTS

- `OfficeIMO.Word.WordTableOfContent`

## RELATED LINKS

- None
