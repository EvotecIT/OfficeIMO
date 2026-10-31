---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Set-OfficePowerPointLayoutPlaceholder
## SYNOPSIS
Updates a template layout placeholder's geometry, text margins and text style.

## SYNTAX
### __AllParameterSets
```powershell
Set-OfficePowerPointLayoutPlaceholder -TemplateLayout <PowerPointLayoutSelector> -Type <PowerPointPlaceholderType> [-Master <int>] [-Index <UInt32>] [-X <OfficeLength>] [-Y <OfficeLength>] [-Width <OfficeLength>] [-Height <OfficeLength>] [-TextMargin <OfficeLength[]>] [-Style <string>] [-Size <OfficeLength>] [-Font <string>] [-Color <OfficeColor>] [-Highlight <OfficeColor>] [-Bold] [-Italic] [-Underline <PowerPointUnderlineStyle>] [-Level <Int32>] [-BulletChar <string>] [-Numbering <PowerPointNumberingScheme>] [-CreateIfMissing] [-PassThru] [-Document <PowerPointDocument>] [<CommonParameters>]
```

## DESCRIPTION
Omitted properties keep their current values. Master and layout positions start at one.
Numbers mean points; text margins use one, two or four values in CSS order.

## EXAMPLES

### EXAMPLE 1
```powershell
Set-OfficePowerPointLayoutPlaceholder -Document $deck -TemplateLayout 'Title and Content' -Type Title -X 2cm -Width 24cm -TextMargin 6pt,8pt -Size 32pt -Bold
```


## PARAMETERS

### -Bold
Gets or sets bold formatting; an explicitly false switch clears it.

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

### -BulletChar
Gets or sets a single bullet character, exclusive with Numbering.

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
Gets or sets text color, resolving theme roles against the selected master.

```yaml
Type: OfficeColor
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
Gets or sets whether to create a missing placeholder.

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

### -Document
Gets or sets an explicit live document; omission uses the enclosing content block.

```yaml
Type: PowerPointDocument
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -Font
Gets or sets the Latin font family.

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

### -Height
Gets or sets the positive placeholder height.

```yaml
Type: OfficeLength
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
Gets or sets the highlight color.

```yaml
Type: OfficeColor
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
Gets or sets the optional native placeholder index.

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
Gets or sets italic formatting; an explicitly false switch clears it.

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

### -Level
Gets or sets the zero-based paragraph indentation level.

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
Gets or sets the one-based master position, defaulting to the first or to the layout metadata's master.

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
Gets or sets the native numbering scheme, exclusive with BulletChar.

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
Gets or sets whether to emit the updated native text box.

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

### -Size
Gets or sets the font size, from one through 4000 points.

```yaml
Type: OfficeLength
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
Gets or sets a native text preset.

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

### -TemplateLayout
Gets or sets a layout name, one-based position, native type or returned layout metadata.

```yaml
Type: PowerPointLayoutSelector
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -TextMargin
Gets or sets margins: all; vertical,horizontal; or top,right,bottom,left.

```yaml
Type: OfficeLength[]
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Type
Gets or sets the placeholder type.

```yaml
Type: PowerPointPlaceholderType
Parameter Sets: __AllParameterSets
Aliases: None
Possible values: Title, Body, CenteredTitle, SubTitle, DateAndTime, SlideNumber, Footer, Header, Object, Chart, Table, ClipArt, Diagram, Media, SlideImage, Picture

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Underline
Gets or sets the native underline style; None removes underlining.

```yaml
Type: PowerPointUnderlineStyle
Parameter Sets: __AllParameterSets
Aliases: None
Possible values: None, Words, Single, Double, Heavy, Dotted, HeavyDotted, Dash, DashHeavy, DashLong, DashLongHeavy, DotDash, DotDashHeavy, DotDotDash, DotDotDashHeavy, Wavy, WavyHeavy, WavyDouble

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Width
Gets or sets the positive placeholder width.

```yaml
Type: OfficeLength
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -X
Gets or sets the left position; negative positions are allowed.

```yaml
Type: OfficeLength
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Y
Gets or sets the top position; negative positions are allowed.

```yaml
Type: OfficeLength
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

- `OfficeIMO.PowerPoint.PowerPointDocument`

## OUTPUTS

- `OfficeIMO.PowerPoint.PowerPointTextBox`

## RELATED LINKS

- None
