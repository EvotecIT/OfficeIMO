---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Add-OfficeWordParagraph
## SYNOPSIS
Adds a paragraph to the current section/header/footer context.

## SYNTAX
### Text (Default)
```powershell
Add-OfficeWordParagraph [[-Text] <string>] [-Document <WordDocument>] [-Alignment <OfficeAlignment>] [-Style <WordParagraphStyles>] [-StyleId <string>] [-PassThru] [<CommonParameters>]
```

### Content
```powershell
Add-OfficeWordParagraph [-Content] <scriptblock> [-Document <WordDocument>] [-Alignment <OfficeAlignment>] [-Style <WordParagraphStyles>] [-StyleId <string>] [-PassThru] [<CommonParameters>]
```

### SectionText
```powershell
Add-OfficeWordParagraph [[-Text] <string>] -Section <WordSection> [-Alignment <OfficeAlignment>] [-Style <WordParagraphStyles>] [-StyleId <string>] [-PassThru] [<CommonParameters>]
```

### SectionContent
```powershell
Add-OfficeWordParagraph [-Content] <scriptblock> -Section <WordSection> [-Alignment <OfficeAlignment>] [-Style <WordParagraphStyles>] [-StyleId <string>] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Acts as the primary DSL container for inline content such as text runs, bold segments, and images.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> Add-OfficeWordParagraph { Add-OfficeWordText -Text 'Hello '; Add-OfficeWordText -Text 'World' -Bold }
```

Outputs “Hello World” with the second word bolded.

### EXAMPLE 2
```powershell
PS> WordParagraph -Text 'Executive summary' -StyleId 'ReportHeading'
```

Applies a paragraph style id, including custom styles already present in a template document.

### EXAMPLE 3
```powershell
PS> $paragraph = $document | Add-OfficeWordParagraph -PassThru
$paragraph | Add-OfficeWordText 'Owner: ' -Bold
$paragraph | Add-OfficeWordText 'Platform'
```

Creates a paragraph on a live document and appends two differently formatted runs.

## PARAMETERS

### -Alignment
Paragraph justification.

```yaml
Type: OfficeAlignment
Parameter Sets: Text, Content, SectionText, SectionContent
Aliases: None
Possible values: Left, Center, Right, Justify, Start, End

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Content
Nested WordText and inline content.

```yaml
Type: ScriptBlock
Parameter Sets: Content, SectionContent
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Document
Live document to receive the paragraph; omitted in a content block.

```yaml
Type: WordDocument
Parameter Sets: Text, Content
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -PassThru
Emit the WordParagraph for further use.

```yaml
Type: SwitchParameter
Parameter Sets: Text, Content, SectionText, SectionContent
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Section
Section to receive the paragraph.

```yaml
Type: WordSection
Parameter Sets: SectionText, SectionContent
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -Style
Paragraph style.

```yaml
Type: WordParagraphStyles
Parameter Sets: Text, Content, SectionText, SectionContent
Aliases: None
Possible values: Normal, Heading1, Heading2, Heading3, Heading4, Heading5, Heading6, Heading7, Heading8, Heading9, ListParagraph, Custom, Title, Subtitle

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -StyleId
Paragraph style id, including custom style ids from a template document.

```yaml
Type: String
Parameter Sets: Text, Content, SectionText, SectionContent
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Text
Initial paragraph text; omission creates an empty paragraph.

```yaml
Type: String
Parameter Sets: Text, SectionText
Aliases: None
Possible values:

Required: False
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `OfficeIMO.Word.WordDocument`
- `OfficeIMO.Word.WordSection`

## OUTPUTS

- `OfficeIMO.Word.WordParagraph`

## RELATED LINKS

- None
