---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Add-OfficeWordTextBox
## SYNOPSIS
Adds an OfficeIMO Word text box to the current Word DSL location.

## SYNTAX
### Default (Default)
```powershell
Add-OfficeWordTextBox [-Text] <string> [-Wrap <WordImageTextWrapping>] [-Width <OfficeLength>] [-Height <OfficeLength>] [-X <OfficeLength>] [-Y <OfficeLength>] [-HorizontalAlignment <WordTextBoxHorizontalAlignment>] [-HorizontalPositionRelativeFrom <WordHorizontalRelativePosition>] [-VerticalPositionRelativeFrom <WordVerticalRelativePosition>] [-AutoFit <WordTextBoxAutoFitType>] [-AutoFitToTextSize] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Adds an OfficeIMO Word text box to the current Word DSL location.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> New-OfficeWord -Path .\Report.docx {
    WordTextBox -Text 'Executive summary' -Width 7cm -Height 2cm -X 1.5cm -Y 1cm -AutoFitToTextSize
}
```

Creates a native OfficeIMO Word text box and applies sizing/positioning through OfficeIMO's text-box API.

## PARAMETERS

### -AutoFit
Explicit OfficeIMO text-box autofit mode.

```yaml
Type: WordTextBoxAutoFitType
Parameter Sets: Default
Aliases: None
Possible values: NoAutoFit, ShrinkTextOnOverflow, ResizeShapeToFitText

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -AutoFitToTextSize
Resize the text box to fit its text.

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

### -Height
Height; numbers mean points and unit suffixes are accepted.

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

### -HorizontalAlignment
Horizontal alignment for anchored text boxes.

```yaml
Type: WordTextBoxHorizontalAlignment
Parameter Sets: Default
Aliases: None
Possible values: Left, Center, Right, Outside, Inside

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -HorizontalPositionRelativeFrom
Horizontal relative position anchor.

```yaml
Type: WordHorizontalRelativePosition
Parameter Sets: Default
Aliases: None
Possible values: Margin, Page, Column, Character, LeftMargin, RightMargin, InsideMargin, OutsideMargin

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PassThru
Emit the created text box.

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

### -Text
Text to place inside the text box.

```yaml
Type: String
Parameter Sets: Default
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -VerticalPositionRelativeFrom
Vertical relative position anchor.

```yaml
Type: WordVerticalRelativePosition
Parameter Sets: Default
Aliases: None
Possible values: Margin, Page, Paragraph, Line, TopMargin, BottomMargin, InsideMargin, OutsideMargin

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Width
Width; numbers mean points and unit suffixes are accepted.

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

### -Wrap
Word text wrapping mode.

```yaml
Type: WordImageTextWrapping
Parameter Sets: Default
Aliases: None
Possible values: InLineWithText, Square, Tight, Through, TopAndBottom, BehindText, InFrontOfText

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -X
Horizontal offset; numbers mean points and unit suffixes are accepted.

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

### -Y
Vertical offset; numbers mean points and unit suffixes are accepted.

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

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `None`

## OUTPUTS

- `OfficeIMO.Word.WordTextBox`

## RELATED LINKS

- None
