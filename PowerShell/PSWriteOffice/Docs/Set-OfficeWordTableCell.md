---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Set-OfficeWordTableCell
## SYNOPSIS
Updates native Word cell text, physical geometry and layout flags.

## SYNTAX
### Default (Default)
```powershell
Set-OfficeWordTableCell [-Cell] <WordTableCell> [-Text <string>] [-MarginTop <OfficeLength>] [-MarginBottom <OfficeLength>] [-MarginLeft <OfficeLength>] [-MarginRight <OfficeLength>] [-BackgroundColor <OfficeColor>] [-ShadingPattern <WordShadingPattern>] [-Width <OfficeLength>] [-TextDirection <WordTextDirection>] [-NoWrap] [-FitText] [-MergeRight <Int32>] [-MergeDown <Int32>] [-SplitHorizontal <Int32>] [-SplitVertical <Int32>] [-CopyParagraphs] [-PassThru] [<CommonParameters>]
```

### WidthPercent
```powershell
Set-OfficeWordTableCell [-Cell] <WordTableCell> -WidthPercent <Double> [-Text <string>] [-MarginTop <OfficeLength>] [-MarginBottom <OfficeLength>] [-MarginLeft <OfficeLength>] [-MarginRight <OfficeLength>] [-BackgroundColor <OfficeColor>] [-ShadingPattern <WordShadingPattern>] [-TextDirection <WordTextDirection>] [-NoWrap] [-FitText] [-MergeRight <Int32>] [-MergeDown <Int32>] [-SplitHorizontal <Int32>] [-SplitVertical <Int32>] [-CopyParagraphs] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Updates native Word cell text, physical geometry and layout flags.

## EXAMPLES

### EXAMPLE 1
```powershell
Set-OfficeWordTableCell -Cell 'Value'
```


### EXAMPLE 2
```powershell
Set-OfficeWordTableCell -Cell 'Value' -WidthPercent 1
```


## PARAMETERS

### -BackgroundColor
BackgroundColor for the native cell.

```yaml
Type: OfficeColor
Parameter Sets: Default, WidthPercent
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Cell
Cell for the native cell.

```yaml
Type: WordTableCell
Parameter Sets: Default, WidthPercent
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -CopyParagraphs
CopyParagraphs for the native cell.

```yaml
Type: SwitchParameter
Parameter Sets: Default, WidthPercent
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -FitText
FitText for the native cell.

```yaml
Type: SwitchParameter
Parameter Sets: Default, WidthPercent
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -MarginBottom
MarginBottom for the native cell.

```yaml
Type: OfficeLength
Parameter Sets: Default, WidthPercent
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -MarginLeft
MarginLeft for the native cell.

```yaml
Type: OfficeLength
Parameter Sets: Default, WidthPercent
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -MarginRight
MarginRight for the native cell.

```yaml
Type: OfficeLength
Parameter Sets: Default, WidthPercent
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -MarginTop
MarginTop for the native cell.

```yaml
Type: OfficeLength
Parameter Sets: Default, WidthPercent
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -MergeDown
MergeDown for the native cell.

```yaml
Type: Int32
Parameter Sets: Default, WidthPercent
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -MergeRight
MergeRight for the native cell.

```yaml
Type: Int32
Parameter Sets: Default, WidthPercent
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NoWrap
NoWrap for the native cell.

```yaml
Type: SwitchParameter
Parameter Sets: Default, WidthPercent
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PassThru
PassThru for the native cell.

```yaml
Type: SwitchParameter
Parameter Sets: Default, WidthPercent
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ShadingPattern
ShadingPattern for the native cell.

```yaml
Type: WordShadingPattern
Parameter Sets: Default, WidthPercent
Aliases: None
Possible values: Nil, Clear, Solid, HorizontalStripe, VerticalStripe, ReverseDiagonalStripe, DiagonalStripe, HorizontalCross, DiagonalCross, ThinHorizontalStripe, ThinVerticalStripe, ThinReverseDiagonalStripe, ThinDiagonalStripe, ThinHorizontalCross, ThinDiagonalCross, Percent5, Percent10, Percent12, Percent15, Percent20, Percent25, Percent30, Percent35, Percent37, Percent40, Percent45, Percent50, Percent55, Percent60, Percent62, Percent65, Percent70, Percent75, Percent80, Percent85, Percent87, Percent90, Percent95

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -SplitHorizontal
SplitHorizontal for the native cell.

```yaml
Type: Int32
Parameter Sets: Default, WidthPercent
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -SplitVertical
SplitVertical for the native cell.

```yaml
Type: Int32
Parameter Sets: Default, WidthPercent
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Text
Text for the native cell.

```yaml
Type: String
Parameter Sets: Default, WidthPercent
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -TextDirection
TextDirection for the native cell.

```yaml
Type: WordTextDirection
Parameter Sets: Default, WidthPercent
Aliases: None
Possible values: LeftToRightTopToBottom, LeftToRightTopToBottom2010, TopToBottomRightToLeft, TopToBottomRightToLeft2010, BottomToTopLeftToRight, BottomToTopLeftToRight2010, LeftToRightTopToBottomRotated, LeftToRightTopToBottomRotated2010, TopToBottomRightToLeftRotated, TopToBottomRightToLeftRotated2010, TopToBottomLeftToRightRotated, TopToBottomLeftToRightRotated2010

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Width
Width for the native cell.

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

### -WidthPercent
WidthPercent for the native cell.

```yaml
Type: Double
Parameter Sets: WidthPercent
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `OfficeIMO.Word.WordTableCell`

## OUTPUTS

- `OfficeIMO.Word.WordTableCell`

## RELATED LINKS

- None
