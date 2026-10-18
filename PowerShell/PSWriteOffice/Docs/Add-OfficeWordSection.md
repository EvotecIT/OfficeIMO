---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Add-OfficeWordSection
## SYNOPSIS
Adds or reuses a section inside the current Word document.

## SYNTAX
### Context (Default)
```powershell
Add-OfficeWordSection [[-Content] <scriptblock>] [-BreakType <WordSectionBreakType>] [-PassThru] [<CommonParameters>]
```

### Document
```powershell
Add-OfficeWordSection [[-Content] <scriptblock>] -Document <WordDocument> [-BreakType <WordSectionBreakType>] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Provides the DSL entry point for section-level operations inside New-OfficeWord.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> New-OfficeWord -Path .\doc.docx { Add-OfficeWordSection { Add-OfficeWordParagraph -Text 'Hello' } }
```

Creates a document and inserts a section that contains a single paragraph.

### EXAMPLE 2
```powershell
PS> $section = $document | Add-OfficeWordSection -BreakType NextPage -PassThru
$section | Add-OfficeWordParagraph -Text 'Appendix' -Style Heading1
```

Adds a section through the document pipeline and uses the returned section as the next explicit target.

## PARAMETERS

### -BreakType
Optional section break type.

```yaml
Type: WordSectionBreakType
Parameter Sets: Context, Document
Aliases: None
Possible values: NextPage, NextColumn, Continuous, EvenPage, OddPage

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Content
DSL scriptblock executed within the section scope.

```yaml
Type: ScriptBlock
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Document
Document that will receive a new section.

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

### -PassThru
Emit the created WordSection.

```yaml
Type: SwitchParameter
Parameter Sets: Context, Document
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

- `OfficeIMO.Word.WordDocument`

## OUTPUTS

- `OfficeIMO.Word.WordSection`

## RELATED LINKS

- None
