---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Add-OfficeWordTabStop
## SYNOPSIS
Adds a tab stop to a Word paragraph.

## SYNTAX
### __AllParameterSets
```powershell
Add-OfficeWordTabStop [-Position] <int> [-Paragraph <WordParagraph>] [-Alignment <WordTabAlignment>] [-Leader <WordTabLeader>] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Thin wrapper over OfficeIMO.Word paragraph tab stops.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> Add-OfficeWordParagraph { Add-OfficeWordTabStop -Position 4320 -Alignment Decimal -Leader Dot }
```

Adds a decimal tab stop at three inches.

## PARAMETERS

### -Alignment
Tab alignment.

```yaml
Type: WordTabAlignment
Parameter Sets: __AllParameterSets
Aliases: None
Possible values: Clear, Left, Start, Center, Right, End, Decimal, Bar, Number

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Leader
Leader character.

```yaml
Type: WordTabLeader
Parameter Sets: __AllParameterSets
Aliases: None
Possible values: None, Dot, Hyphen, Underscore, Heavy, MiddleDot

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Paragraph
Paragraph to update.

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
Emit the created tab stop.

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

### -Position
Tab position in twips.

```yaml
Type: Int32
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

- `OfficeIMO.Word.WordTabStop`

## RELATED LINKS

- None
