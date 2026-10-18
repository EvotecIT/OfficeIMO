---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Set-OfficeWordTextStyle
## SYNOPSIS
Updates styling on Word text.

## SYNTAX
### __AllParameterSets
```powershell
Set-OfficeWordTextStyle [[-InputObject] <WordParagraph>] [-Text <string>] [-Style <WordCharacterStyles>] [-StyleId <string>] [-Bold <Boolean>] [-Italic <Boolean>] [-Underline <WordUnderlineStyle>] [-Color <string>] [-FontSize <Double>] [-FontFamily <string>] [-Highlight <WordHighlightColor>] [-Strike <Boolean>] [-DoubleStrike <Boolean>] [-CapsStyle <WordCapsStyle>] [-Spacing <Int32>] [-Outline <Boolean>] [-Shadow <Boolean>] [-Emboss <Boolean>] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Updates styling on Word text.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> WordParagraph -Text 'Warning' -PassThru | Set-OfficeWordTextStyle -Bold $true -Color '#C00000'
```

Applies bold red styling to matching text.

## PARAMETERS

### -Bold
Set or clear bold formatting.

```yaml
Type: Boolean
Parameter Sets: __AllParameterSets
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
Parameter Sets: __AllParameterSets
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
Type: String
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -DoubleStrike
Set or clear double strikethrough.

```yaml
Type: Boolean
Parameter Sets: __AllParameterSets
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
Type: Boolean
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -FontFamily
Font family name.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -FontSize
Font size in points.

```yaml
Type: Double
Parameter Sets: __AllParameterSets
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
Parameter Sets: __AllParameterSets
Aliases: None
Possible values: Black, Blue, Cyan, Green, Magenta, Red, Yellow, White, DarkBlue, DarkCyan, DarkGreen, DarkMagenta, DarkRed, DarkYellow, DarkGray, LightGray, None

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -InputObject
Word text item to update.

```yaml
Type: WordParagraph
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: 0
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -Italic
Set or clear italic formatting.

```yaml
Type: Boolean
Parameter Sets: __AllParameterSets
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
Type: Boolean
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PassThru
Emit the updated text item.

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

### -Shadow
Set or clear shadow effect.

```yaml
Type: Boolean
Parameter Sets: __AllParameterSets
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
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Strike
Set or clear strikethrough.

```yaml
Type: Boolean
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
Character style to apply.

```yaml
Type: WordCharacterStyles
Parameter Sets: __AllParameterSets
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
Replace the text.

```yaml
Type: String
Parameter Sets: __AllParameterSets
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
Parameter Sets: __AllParameterSets
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
