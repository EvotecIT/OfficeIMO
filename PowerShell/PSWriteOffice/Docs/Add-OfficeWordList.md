---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Add-OfficeWordList
## SYNOPSIS
Starts a list inside the current section or paragraph anchor.

## SYNTAX
### __AllParameterSets
```powershell
Add-OfficeWordList [[-Content] <scriptblock>] [[-Style] <WordListStyle>] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Creates a temporary anchor paragraph, spawns an OfficeIMO list, and runs child WordListItem commands.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> Add-OfficeWordList -Style 'Numbered' { Add-OfficeWordListItem -Text 'Plan'; Add-OfficeWordListItem -Text 'Execute' }
```

Creates a numbered list with two steps.

## PARAMETERS

### -Content
Scriptblock executed within the list scope.

```yaml
Type: ScriptBlock
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PassThru
Emit the object created or changed by the command.

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
Built-in list style or custom numbering scheme.

```yaml
Type: WordListStyle
Parameter Sets: __AllParameterSets
Aliases: Type
Possible values: Bulleted, ArticleSections, Headings111, HeadingIA1, Chapters, BulletedChars, Heading1ai, Headings111Shifted, LowerLetterWithBracket, LowerLetterWithDot, UpperLetterWithDot, UpperLetterWithBracket, Custom, Numbered

Required: False
Position: 1
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `None`

## OUTPUTS

- `None`

## RELATED LINKS

- None
