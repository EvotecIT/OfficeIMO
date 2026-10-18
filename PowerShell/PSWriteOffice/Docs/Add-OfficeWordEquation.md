---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Add-OfficeWordEquation
## SYNOPSIS
Adds an Office Math equation to a Word document or paragraph.

## SYNTAX
### __AllParameterSets
```powershell
Add-OfficeWordEquation [-Omml] <string> [-Paragraph <WordParagraph>] [-Document <WordDocument>] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Accepts OMML and keeps conversion/parsing outside the cmdlet.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $omml = '<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><m:r><m:t>x+1</m:t></m:r></m:oMath>'
New-OfficeWord -Path .\Formula.docx {
    Add-OfficeWordParagraph -Text 'The following expression is stored as Office Math.'
    Add-OfficeWordEquation -Omml $omml
}
```

Inserts prebuilt OMML into the current document; conversion to OMML is intentionally outside the cmdlet.

## PARAMETERS

### -Document
Document to receive a new equation paragraph.

```yaml
Type: WordDocument
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Omml
Office Math Markup Language content.

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

### -Paragraph
Paragraph to receive the equation.

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
Emit the equation paragraph.

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

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `OfficeIMO.Word.WordParagraph`

## OUTPUTS

- `OfficeIMO.Word.WordParagraph`

## RELATED LINKS

- None
