---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# New-OfficePdfTableCell
## SYNOPSIS
Creates a reusable PDF table cell definition for explicit table rows.

## SYNTAX
### __AllParameterSets
```powershell
New-OfficePdfTableCell [[-Text] <string>] [-Run <Object[]>] [-ColumnSpan <int>] [-RowSpan <int>] [-TextColor <string>] [-FillColor <string>] [-FontSize <Double>] [-Bold] [-Italic] [-Underline] [-UnderlineStyle <string>] [-Strike] [-Align <PdfColumnAlign>] [-VerticalAlign <PdfCellVerticalAlign>] [-CheckBox <PdfTableCellCheckBox[]>] [-Image <PdfTableCellImage[]>] [-FormField <PdfTableCellFormField[]>] [-LinkUri <string>] [-LinkDestinationName <string>] [-LinkContents <string>] [-NamedDestinationName <string>] [-NoWrap] [<CommonParameters>]
```

## DESCRIPTION
Creates a reusable PDF table cell definition for explicit table rows.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $row = @(New-OfficePdfTableCell -Text 'Identity systems' -ColumnSpan 3 -FillColor '#DBEAFE' -TextColor '#1E3A8A' -Bold)
```

The returned cell can be passed to PdfTable inside explicit row arrays.

## PARAMETERS

### -Align
Horizontal cell alignment.

```yaml
Type: PdfColumnAlign
Parameter Sets: __AllParameterSets
Aliases: None
Possible values: Left, Center, Right

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

### -CheckBox
Typed check boxes rendered inside the cell.

```yaml
Type: PdfTableCellCheckBox[]
Parameter Sets: __AllParameterSets
Aliases: CheckBoxes
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
Cell font size in PDF points.

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

### -FormField
Typed text or choice form fields rendered inside the cell.

```yaml
Type: PdfTableCellFormField[]
Parameter Sets: __AllParameterSets
Aliases: FormFields
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Image
Typed images rendered inside the cell.

```yaml
Type: PdfTableCellImage[]
Parameter Sets: __AllParameterSets
Aliases: Images
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

### -LinkContents
Accessible annotation text for the cell link.

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

### -LinkDestinationName
Named PDF destination linked from the cell.

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

### -LinkUri
Absolute or catalog-base-relative URI linked from the cell.

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

### -NamedDestinationName
Named PDF destination defined at this cell.

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

### -NoWrap
Keep the cell content on one visual line.

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
Rich text runs for the cell. Each run can be created with TextRun/PdfTextRun or provided as a hashtable/object.

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
Optional underline style name. PDF table rendering treats any supported value as underline.

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

### -VerticalAlign
Vertical cell alignment.

```yaml
Type: PdfCellVerticalAlign
Parameter Sets: __AllParameterSets
Aliases: None
Possible values: Top, Middle, Bottom

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
