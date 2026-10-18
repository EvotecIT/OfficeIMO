---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Add-OfficeWordBreak
## SYNOPSIS
Adds a break to a Word paragraph.

## SYNTAX
### __AllParameterSets
```powershell
Add-OfficeWordBreak [-Paragraph <WordParagraph>] [-BreakType <WordBreakType>] [-Count <int>] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
By default this creates a soft line break, equivalent to Shift+Enter in Word.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> Add-OfficeWordParagraph { Add-OfficeWordText 'Line 1'; Add-OfficeWordBreak; Add-OfficeWordText 'Line 2' }
```

Writes both lines in the same paragraph separated by a soft break.

## PARAMETERS

### -BreakType
Optional break type, for example Page or Column.

```yaml
Type: WordBreakType
Parameter Sets: __AllParameterSets
Aliases: None
Possible values: Page, Column, TextWrapping

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Count
Number of breaks to add.

```yaml
Type: Int32
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Paragraph
Target paragraph. When omitted, the current DSL paragraph is used or created.

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
Emit the paragraph returned by the final break for additional native OfficeIMO chaining.

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
