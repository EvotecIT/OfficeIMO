---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Add-OfficeWordTableOfContents
## SYNOPSIS
Adds a table of contents to a Word document.

## SYNTAX
### __AllParameterSets
```powershell
Add-OfficeWordTableOfContents [-Document <WordDocument>] [-Style <WordTableOfContentsStyle>] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Adds a table of contents to a Word document.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> New-OfficeWord -Path .\ExecutiveReport.docx {
    Add-OfficeWordTableOfContents -Style Template1
    Add-OfficeWordParagraph -Text 'Executive summary' -Style Heading1
    Add-OfficeWordParagraph -Text 'Summary text'
    Add-OfficeWordParagraph -Text 'Appendix' -Style Heading1
    Add-OfficeWordParagraph -Text 'Supporting details'
    Update-OfficeWordTableOfContents
}
```

Creates a navigable report outline and marks the TOC for refresh when the document opens.

## PARAMETERS

### -Document
Document to modify when provided explicitly.

```yaml
Type: WordDocument
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -PassThru
Emit the created table of contents.

```yaml
Type: SwitchParameter
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Style
Table of contents template style.

```yaml
Type: WordTableOfContentsStyle
Parameter Sets: __AllParameterSets
Aliases: None
Possible values: Template1, Template2

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `OfficeIMO.Word.WordDocument`

## OUTPUTS

- `OfficeIMO.Word.WordTableOfContent`

## RELATED LINKS

- None
