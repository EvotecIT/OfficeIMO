---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# New-OfficeWordTableCell
## SYNOPSIS
Creates a reusable Word table cell definition for explicit table rows.

## SYNTAX
### __AllParameterSets
```powershell
New-OfficeWordTableCell [[-Text] <string>] [-Run <Object[]>] [-ColumnSpan <int>] [-RowSpan <int>] [-TextColor <string>] [-FillColor <string>] [-FontSize <Double>] [-Bold] [-Italic] [-Underline] [-UnderlineStyle <WordUnderlineStyle>] [-Strike] [-Align <WordParagraphAlignment>] [-VerticalAlign <WordTableVerticalAlignment>] [<CommonParameters>]
```

## DESCRIPTION
Creates a reusable Word table cell definition for explicit table rows.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $row = @(New-OfficeWordTableCell -Text 'Identity systems' -ColumnSpan 3)
```

The returned cell can be passed to WordTable inside explicit row arrays.

## PARAMETERS

### -Align
Horizontal cell alignment.

```yaml
Type: WordParagraphAlignment
Parameter Sets: __AllParameterSets
Aliases: None
Possible values: Left, Start, Center, Right, End, Both, MediumKashida, Distribute, NumTab, HighKashida, LowKashida, ThaiDistribute

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Bold
Render the cell text in bold.

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

### -ColumnSpan
Number of logical columns covered by the cell.

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

### -FillColor
Cell fill color. Named colors and hexadecimal colors are accepted.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: BackgroundColor, CellFill
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -FontSize
Cell font size in points.

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

### -Italic
Render the cell text in italics.

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

### -RowSpan
Number of logical rows covered by the cell.

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

### -Run
Rich text runs for the cell. Each run can be created with TextRun/WordTextRun or provided as a hashtable/object.

```yaml
Type: Object[]
Parameter Sets: __AllParameterSets
Aliases: Runs
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Strike
Render the cell text with strikethrough.

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

### -Text
Cell text.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -TextColor
Cell text color. Named colors and hexadecimal colors are accepted.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: Color, FontColor
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Underline
Render the cell text with underline.

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

### -UnderlineStyle
Optional Word underline style.

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

### -VerticalAlign
Vertical cell alignment.

```yaml
Type: WordTableVerticalAlignment
Parameter Sets: __AllParameterSets
Aliases: None
Possible values: Top, Center, Bottom

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

- `PSWriteOffice.Services.Table.OfficeTableCellSpec`: Describes a logical table cell that can be rendered by multiple Office table surfaces.

## RELATED LINKS

- None
