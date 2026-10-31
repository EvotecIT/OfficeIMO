---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Set-OfficeWordTableCell
## SYNOPSIS
Updates OfficeIMO Word table-cell content, layout, and merge settings.

## SYNTAX
### __AllParameterSets
```powershell
Set-OfficeWordTableCell [-Cell] <WordTableCell> [-Text <string>] [-MarginTopCentimeters <Double>] [-MarginBottomCentimeters <Double>] [-MarginLeftCentimeters <Double>] [-MarginRightCentimeters <Double>] [-ShadingFillColor <string>] [-ShadingPattern <WordShadingPattern>] [-Width <Int32>] [-WidthType <WordTableWidthUnit>] [-TextDirection <WordTextDirection>] [-WrapText <Boolean>] [-FitText <Boolean>] [-MergeRight <Int32>] [-MergeDown <Int32>] [-SplitHorizontal <Int32>] [-SplitVertical <Int32>] [-CopyParagraphs] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Updates OfficeIMO Word table-cell content, layout, and merge settings.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $doc = Open-OfficeWord -Path .\Handover.docx
$table = Find-OfficeWordTable -Document $doc -Text 'Risk marker' | Select-Object -First 1
$table |
    Get-OfficeWordTableCell -Row 2 -Column 2 |
    Set-OfficeWordTableCell -Text 'Investigating' -ShadingFillColor '#fff2cc' -ShadingPattern Clear
$doc | Close-OfficeWord -Save
```

Finds an existing table by text, replaces a target cell value, applies shading, and saves the document.

### EXAMPLE 2
```powershell
PS> $doc = Open-OfficeWord -Path .\Report.docx
$table = $doc | Get-OfficeWordTable | Select-Object -First 1
$table |
    Get-OfficeWordTableCell -Column 2 |
    Set-OfficeWordTableCell -ShadingFillColor '#fff1f0' -ShadingPattern Clear -Width 2400 -WidthType Dxa
$doc | Save-OfficeWord -Path .\Report-StatusCells.docx
```

Reads cells from an OfficeIMO table object, applies cell shading and width, and saves the updated document.

### EXAMPLE 3
```powershell
PS> $table | Get-OfficeWordTableCell | Set-OfficeWordTableCell -MarginTopCentimeters 0.12 -MarginBottomCentimeters 0.12 -MarginLeftCentimeters 0.16 -MarginRightCentimeters 0.16
```

Sets independent cell margins in centimeters. Omitted margins retain their values; pass null to remove an explicit margin.

### EXAMPLE 4
```powershell
PS> $doc = Open-OfficeWord -Path .\Report.docx
$table = $doc | Get-OfficeWordTable | Select-Object -First 1
$table |
    Get-OfficeWordTableCell -Row 0 -Column 0 |
    Set-OfficeWordTableCell -MergeRight 2 -CopyParagraphs
$doc | Save-OfficeWord -Path .\Report-MergedHeader.docx
```

Uses the OfficeIMO merge operation exposed by the thin table-cell wrapper.

## PARAMETERS

### -Cell
Table cell to update.

```yaml
Type: WordTableCell
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -CopyParagraphs
Copy paragraphs while merging cells.

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

### -FitText
Whether text should fit within the cell.

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

### -MarginBottomCentimeters
Bottom cell padding in centimeters. Pass null to remove the explicit margin.

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

### -MarginLeftCentimeters
Left cell padding in centimeters. Pass null to remove the explicit margin.

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

### -MarginRightCentimeters
Right cell padding in centimeters. Pass null to remove the explicit margin.

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

### -MarginTopCentimeters
Top cell padding in centimeters. Pass null to remove the explicit margin.

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

### -MergeDown
Number of cells to merge downward.

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

### -MergeRight
Number of cells to merge to the right.

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

### -PassThru
Emit the updated table cell.

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

### -ShadingFillColor
Cell shading fill color. Named colors and hexadecimal values are accepted.

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

### -ShadingPattern
Cell shading pattern.

```yaml
Type: WordShadingPattern
Parameter Sets: __AllParameterSets
Aliases: None
Possible values: Nil, Clear, Solid, HorizontalStripe, VerticalStripe, ReverseDiagonalStripe, DiagonalStripe, HorizontalCross, DiagonalCross, ThinHorizontalStripe, ThinVerticalStripe, ThinReverseDiagonalStripe, ThinDiagonalStripe, ThinHorizontalCross, ThinDiagonalCross, Percent5, Percent10, Percent12, Percent15, Percent20, Percent25, Percent30, Percent35, Percent37, Percent40, Percent45, Percent50, Percent55, Percent60, Percent62, Percent65, Percent70, Percent75, Percent80, Percent85, Percent87, Percent90, Percent95

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -SplitHorizontal
Number of columns to split the cell into.

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

### -SplitVertical
Number of rows to split the cell into.

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

### -Text
Replace the visible cell text.

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

### -TextDirection
Cell text direction.

```yaml
Type: WordTextDirection
Parameter Sets: __AllParameterSets
Aliases: None
Possible values: LeftToRightTopToBottom, LeftToRightTopToBottom2010, TopToBottomRightToLeft, TopToBottomRightToLeft2010, BottomToTopLeftToRight, BottomToTopLeftToRight2010, LeftToRightTopToBottomRotated, LeftToRightTopToBottomRotated2010, TopToBottomRightToLeftRotated, TopToBottomRightToLeftRotated2010, TopToBottomLeftToRightRotated, TopToBottomLeftToRightRotated2010

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Width
Native cell preference: twips for Dxa or fiftieths of a percent for Pct. Null removes the preference.

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

### -WidthType
Cell width unit type.

```yaml
Type: WordTableWidthUnit
Parameter Sets: __AllParameterSets
Aliases: None
Possible values: Nil, Pct, Dxa, Auto

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -WrapText
Whether text wraps in the cell.

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

- `OfficeIMO.Word.WordTableCell`

## OUTPUTS

- `OfficeIMO.Word.WordTableCell`

## RELATED LINKS

- None
