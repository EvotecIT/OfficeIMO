---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Add-OfficeWordEndnote
## SYNOPSIS
Adds an endnote reference to a Word paragraph.

## SYNTAX
### __AllParameterSets
```powershell
Add-OfficeWordEndnote [-Text] <string> [-Paragraph <WordParagraph>] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Adds an endnote reference to a Word paragraph.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> Add-OfficeWordParagraph -Text 'Appendix reference' { Add-OfficeWordEndnote -Text 'Full calculation appears in the appendix.' }
```

Creates an endnote reference on the current paragraph.

## PARAMETERS

### -Paragraph
Paragraph to receive the endnote reference.

```yaml
Type: WordParagraph
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
Emit the created endnote paragraph.

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

### -Text
Endnote text.

```yaml
Type: String
Parameter Sets: __AllParameterSets
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

- `OfficeIMO.Word.WordParagraph`

## OUTPUTS

- `OfficeIMO.Word.WordParagraph`

## RELATED LINKS

- None
