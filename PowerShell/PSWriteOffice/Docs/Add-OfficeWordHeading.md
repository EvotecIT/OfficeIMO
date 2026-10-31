---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Add-OfficeWordHeading
## SYNOPSIS
Adds a heading using the document's built-in heading, title or subtitle style.

## SYNTAX
### Text (Default)
```powershell
Add-OfficeWordHeading [-Text] <string> [-Level <WordHeadingLevel>] [-Alignment <OfficeAlignment>] [-Document <WordDocument>] [-PassThru] [<CommonParameters>]
```

### Content
```powershell
Add-OfficeWordHeading [-Content] <scriptblock> [-Level <WordHeadingLevel>] [-Alignment <OfficeAlignment>] [-Document <WordDocument>] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Levels 1–9 participate in document navigation and tables of contents. A content block accepts WordText runs.

## EXAMPLES

### EXAMPLE 1
```powershell
WordHeading 'Service review'
            WordHeading 'Availability' -Level 2
```


### EXAMPLE 2
```powershell
Add-OfficeWordHeading -Document $document -Level 2 { WordText 'Open incidents '; WordText '(3)' -Bold }
```


## PARAMETERS

### -Alignment
Optional horizontal alignment.

```yaml
Type: OfficeAlignment
Parameter Sets: Text, Content
Aliases: None
Possible values: Left, Center, Right, Justify, Start, End

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Content
Rich heading content composed with WordText.

```yaml
Type: ScriptBlock
Parameter Sets: Content
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Document
Live document; omitted inside a Word content block.

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

### -Level
Heading level 1–9, Title or Subtitle. Defaults to Heading1.

```yaml
Type: WordHeadingLevel
Parameter Sets: Text, Content
Aliases: None
Possible values: Heading1, Heading2, Heading3, Heading4, Heading5, Heading6, Heading7, Heading8, Heading9, Title, Subtitle

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PassThru
Emit the object created or changed by the command.

```yaml
Type: SwitchParameter
Parameter Sets: Text, Content
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Text
Heading text.

```yaml
Type: String
Parameter Sets: Text
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `OfficeIMO.Word.WordDocument`

## OUTPUTS

- `OfficeIMO.Word.WordParagraph`

## RELATED LINKS

- None
