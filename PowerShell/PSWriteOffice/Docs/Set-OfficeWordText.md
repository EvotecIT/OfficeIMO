---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Set-OfficeWordText
## SYNOPSIS
Updates styling on Word text.

## SYNTAX
### Default (Default)
```powershell
Set-OfficeWordText [[-Paragraph] <WordParagraph>] [-Text <string>] [-Style <WordCharacterStyles>] [-StyleId <string>] [-Bold] [-Italic] [-Underline <WordUnderlineStyle>] [-Color <OfficeColor>] [-Size <OfficeLength>] [-Font <string>] [-Highlight <WordHighlightColor>] [-Strikethrough] [-DoubleStrikethrough] [-CapsStyle <WordCapsStyle>] [-Spacing <Int32>] [-Outline] [-Shadow] [-Emboss] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Updates styling on Word text.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> WordParagraph -Text 'Warning' -PassThru | Set-OfficeWordText -Bold -Color '#C00000'
```

Applies bold red styling to the supplied paragraph or scoped run.

## PARAMETERS

### -Bold
Set or clear bold formatting.

```yaml
Type: SwitchParameter
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -CapsStyle
Capitalization style.

```yaml
Type: WordCapsStyle
Parameter Sets: Default
Aliases: None
Possible values: None, Caps, SmallCaps

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Color
Text color. Named colors and hexadecimal values are accepted.

```yaml
Type: OfficeColor
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -DoubleStrikethrough
Set or clear double strikethrough.

```yaml
Type: SwitchParameter
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Emboss
Set or clear emboss effect.

```yaml
Type: SwitchParameter
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Font
Font family name.

```yaml
Type: String
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Highlight
Highlight color.

```yaml
Type: WordHighlightColor
Parameter Sets: Default
Aliases: None
Possible values: Black, Blue, Cyan, Green, Magenta, Red, Yellow, White, DarkBlue, DarkCyan, DarkGreen, DarkMagenta, DarkRed, DarkYellow, DarkGray, LightGray, None

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Italic
Set or clear italic formatting.

```yaml
Type: SwitchParameter
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Outline
Set or clear outline effect.

```yaml
Type: SwitchParameter
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Paragraph
Word text item to update.

```yaml
Type: WordParagraph
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: 0
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -PassThru
Emit the updated text item.

```yaml
Type: SwitchParameter
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Shadow
Set or clear shadow effect.

```yaml
Type: SwitchParameter
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Size
Font size in points.

```yaml
Type: OfficeLength
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Spacing
Character spacing in twentieths of a point.

```yaml
Type: Int32
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Strikethrough
Set or clear strikethrough.

```yaml
Type: SwitchParameter
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Style
Character style to apply.

```yaml
Type: WordCharacterStyles
Parameter Sets: Default
Aliases: None
Possible values: DefaultParagraphFont, Heading1Char, Heading2Char, Heading3Char, Heading4Char, Heading5Char, Heading6Char, Heading7Char, Heading8Char, Heading9Char

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -StyleId
Character style id to apply.

```yaml
Type: String
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Text
Replace the text.

```yaml
Type: String
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Underline
Underline style to apply.

```yaml
Type: WordUnderlineStyle
Parameter Sets: Default
Aliases: None
Possible values: Single, Words, Double, Thick, Dotted, DottedHeavy, Dash, DashedHeavy, DashLong, DashLongHeavy, DotDash, DashDotHeavy, DotDotDash, DashDotDotHeavy, Wave, WavyHeavy, WavyDouble, None

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
