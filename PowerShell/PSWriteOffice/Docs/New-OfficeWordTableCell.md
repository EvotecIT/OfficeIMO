---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# New-OfficeWordTableCell
## SYNOPSIS
Creates a plain or recorded rich Word cell for explicit table rows.

## SYNTAX
### Text (Default)
```powershell
New-OfficeWordTableCell [[-Text] <string>] [-ColumnSpan <int>] [-RowSpan <int>] [-BackgroundColor <OfficeColor>] [-Alignment <OfficeAlignment>] [-VerticalAlignment <WordTableVerticalAlignment>] [-Bold] [-Italic] [-Strikethrough] [-Underline <WordUnderlineStyle>] [-Color <OfficeColor>] [-Font <string>] [-Size <OfficeLength>] [<CommonParameters>]
```

### Content
```powershell
New-OfficeWordTableCell [-Content] <scriptblock> [-ColumnSpan <int>] [-RowSpan <int>] [-BackgroundColor <OfficeColor>] [-Alignment <OfficeAlignment>] [-VerticalAlignment <WordTableVerticalAlignment>] [<CommonParameters>]
```

## DESCRIPTION
Creates a plain or recorded rich Word cell for explicit table rows.

## EXAMPLES

### EXAMPLE 1
```powershell
New-OfficeWordTableCell -Alignment 'Value'
```


### EXAMPLE 2
```powershell
New-OfficeWordTableCell -Content { }
```


## PARAMETERS

### -Alignment
Alignment for the cell.

```yaml
Type: OfficeAlignment
Parameter Sets: Text, Content
Aliases: None
Possible values: Left, Center, Right, Justify, Start, End

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -BackgroundColor
BackgroundColor for the cell.

```yaml
Type: OfficeColor
Parameter Sets: Text, Content
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Bold
Bold for the cell.

```yaml
Type: SwitchParameter
Parameter Sets: Text
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Color
Color for the cell.

```yaml
Type: OfficeColor
Parameter Sets: Text
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ColumnSpan
ColumnSpan for the cell.

```yaml
Type: Int32
Parameter Sets: Text, Content
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Content
Content for the cell.

```yaml
Type: ScriptBlock
Parameter Sets: Content
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Font
Font for the cell.

```yaml
Type: String
Parameter Sets: Text
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Italic
Italic for the cell.

```yaml
Type: SwitchParameter
Parameter Sets: Text
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -RowSpan
RowSpan for the cell.

```yaml
Type: Int32
Parameter Sets: Text, Content
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Size
Size for the cell.

```yaml
Type: OfficeLength
Parameter Sets: Text
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Strikethrough
Strikethrough for the cell.

```yaml
Type: SwitchParameter
Parameter Sets: Text
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Text
Text for the cell.

```yaml
Type: String
Parameter Sets: Text
Aliases: None
Possible values:

Required: False
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Underline
Underline for the cell.

```yaml
Type: WordUnderlineStyle
Parameter Sets: Text
Aliases: None
Possible values: Single, Words, Double, Thick, Dotted, DottedHeavy, Dash, DashedHeavy, DashLong, DashLongHeavy, DotDash, DashDotHeavy, DotDotDash, DashDotDotHeavy, Wave, WavyHeavy, WavyDouble, None

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -VerticalAlignment
VerticalAlignment for the cell.

```yaml
Type: WordTableVerticalAlignment
Parameter Sets: Text, Content
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
