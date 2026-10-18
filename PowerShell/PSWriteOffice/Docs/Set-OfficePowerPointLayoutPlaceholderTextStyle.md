---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Set-OfficePowerPointLayoutPlaceholderTextStyle
## SYNOPSIS
Sets layout placeholder text style and bullet/numbering settings.

## SYNTAX
### __AllParameterSets
```powershell
Set-OfficePowerPointLayoutPlaceholderTextStyle -Layout <int> -PlaceholderType <PowerPointPlaceholderType> [-Presentation <PowerPointPresentation>] [-Master <int>] [-Index <UInt32>] [-Style <string>] [-FontSize <Int32>] [-FontName <string>] [-Color <string>] [-Bold <Boolean>] [-Italic <Boolean>] [-Underline <Boolean>] [-HighlightColor <string>] [-Level <Int32>] [-BulletChar <string>] [-Numbering <PowerPointNumberingScheme>] [-CreateIfMissing] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Sets layout placeholder text style and bullet/numbering settings.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> Set-OfficePowerPointLayoutPlaceholderTextStyle -Presentation $ppt -Master 0 -Layout 1 -PlaceholderType Title -Style Title
```

Applies the Title preset to the layout placeholder.

### EXAMPLE 2
```powershell
PS> New-OfficePowerPoint -Path .\deck.pptx {
  $layout = Get-OfficePowerPointLayout | Select-Object -First 1
  Set-OfficePowerPointLayoutPlaceholderTextStyle -Master $layout.MasterIndex -Layout $layout.LayoutIndex -PlaceholderType Title -Style Title -FontSize 36 -Bold $true
}
```

Uses the DSL context to resolve the presentation.

## PARAMETERS

### -Bold
Apply bold formatting.

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

### -BulletChar
Optional bullet character (ignored when -Numbering is supplied).

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

### -CreateIfMissing
Create the placeholder if it is missing.

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

### -FontName
Font name (Latin).

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

### -HighlightColor
Highlight color. Named colors and hexadecimal values are accepted.

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

### -Index
Optional placeholder index.

```yaml
Type: UInt32
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Italic
Apply italic formatting.

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

### -Layout
Layout index within the master.

```yaml
Type: Int32
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Level
Paragraph level (0-8) to set before applying style.

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

### -Master
Slide master index.

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

### -Numbering
Optional numbering scheme name (e.g. ArabicPeriod, RomanUpper).

```yaml
Type: PowerPointNumberingScheme
Parameter Sets: __AllParameterSets
Aliases: None
Possible values: AlphaLowerCharacterParenBoth, AlphaUpperCharacterParenBoth, AlphaLowerCharacterParenR, AlphaUpperCharacterParenR, AlphaLowerCharacterPeriod, AlphaUpperCharacterPeriod, ArabicParenBoth, ArabicParenR, ArabicPeriod, ArabicPlain, RomanLowerCharacterParenBoth, RomanUpperCharacterParenBoth, RomanLowerCharacterParenR, RomanUpperCharacterParenR, RomanLowerCharacterPeriod, RomanUpperCharacterPeriod, CircleNumberDoubleBytePlain, CircleNumberWingdingsBlackPlain, CircleNumberWingdingsWhitePlain, ArabicDoubleBytePeriod, ArabicDoubleBytePlain, EastAsianSimplifiedChinesePeriod, EastAsianSimplifiedChinesePlain, EastAsianTraditionalChinesePeriod, EastAsianTraditionalChinesePlain, EastAsianJapaneseDoubleBytePeriod, EastAsianJapaneseKoreanPlain, EastAsianJapaneseKoreanPeriod, Arabic1Minus, Arabic2Minus, Hebrew2Minus, ThaiAlphaPeriod, ThaiAlphaParenthesisRight, ThaiAlphaParenthesisBoth, ThaiNumberPeriod, ThaiNumberParenthesisRight, ThaiNumberParenthesisBoth, HindiAlphaPeriod, HindiNumPeriod, HindiNumberParenthesisRight, HindiAlpha1Period

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PassThru
Emit the placeholder textbox after update.

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

### -PlaceholderType
Placeholder type to target.

```yaml
Type: PowerPointPlaceholderType
Parameter Sets: __AllParameterSets
Aliases: Type
Possible values: Title, Body, CenteredTitle, SubTitle, DateAndTime, SlideNumber, Footer, Header, Object, Chart, Table, ClipArt, Diagram, Media, SlideImage, Picture

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Presentation
Presentation to update (optional inside DSL).

```yaml
Type: PowerPointPresentation
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -Style
Named style preset (Title, Subtitle, Body, Caption, Emphasis).

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: None
Possible values: Title, Subtitle, Body, Caption, Emphasis

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Underline
Apply underline formatting.

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

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `OfficeIMO.PowerPoint.PowerPointPresentation`

## OUTPUTS

- `OfficeIMO.PowerPoint.PowerPointTextBox`

## RELATED LINKS

- None
