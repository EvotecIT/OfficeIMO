---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Add-OfficePdfTable
## SYNOPSIS
Adds a table to a PDF document.

## SYNTAX
### Context (Default)
```powershell
Add-OfficePdfTable [-InputObject] <Object[]> [-Property <string[]>] [-Header <string[]>] [-View <OfficeTableView>] [-CollectionSeparator <string>] [-DictionaryEntrySeparator <string>] [-DictionaryKeyValueSeparator <string>] [-MaxCollectionItems <int>] [-MaxNestingDepth <int>] [-ColumnFormat <hashtable>] [-Alignment <OfficeAlignment>] [-TableStyle <string>] [-HeaderBackgroundColor <OfficeColor>] [-HeaderColor <OfficeColor>] [-Color <OfficeColor>] [-BandColor <OfficeColor>] [-BorderColor <OfficeColor>] [-BorderWidth <OfficeLength>] [-Size <OfficeLength>] [-HeaderSize <OfficeLength>] [-LineHeight <Double>] [-CellPaddingX <OfficeLength>] [-CellPaddingY <OfficeLength>] [-SpacingBefore <OfficeLength>] [-SpacingAfter <OfficeLength>] [-Caption <string>] [-CaptionAlignment <OfficeAlignment>] [-CaptionColor <OfficeColor>] [-CaptionSize <OfficeLength>] [-ColumnWidth <OfficeLength[]>] [-ColumnWeight <double[]>] [-ColumnAlignment <OfficeAlignment[]>] [-AutoFit] [-NoNumericAlignment] [-ShrinkToFit] [-MinimumSize <OfficeLength>] [-KeepTogether] [-KeepWithNext] [-NoBorder] [-NoHeaderBackground] [-NoBanding] [-HeaderRowCount <Int32>] [-RepeatHeaderRowCount <Int32>] [-FooterRowCount <Int32>] [<CommonParameters>]
```

### Document
```powershell
Add-OfficePdfTable [-InputObject] <Object[]> -Document <PdfDocument> [-Property <string[]>] [-Header <string[]>] [-View <OfficeTableView>] [-CollectionSeparator <string>] [-DictionaryEntrySeparator <string>] [-DictionaryKeyValueSeparator <string>] [-MaxCollectionItems <int>] [-MaxNestingDepth <int>] [-ColumnFormat <hashtable>] [-Alignment <OfficeAlignment>] [-TableStyle <string>] [-HeaderBackgroundColor <OfficeColor>] [-HeaderColor <OfficeColor>] [-Color <OfficeColor>] [-BandColor <OfficeColor>] [-BorderColor <OfficeColor>] [-BorderWidth <OfficeLength>] [-Size <OfficeLength>] [-HeaderSize <OfficeLength>] [-LineHeight <Double>] [-CellPaddingX <OfficeLength>] [-CellPaddingY <OfficeLength>] [-SpacingBefore <OfficeLength>] [-SpacingAfter <OfficeLength>] [-Caption <string>] [-CaptionAlignment <OfficeAlignment>] [-CaptionColor <OfficeColor>] [-CaptionSize <OfficeLength>] [-ColumnWidth <OfficeLength[]>] [-ColumnWeight <double[]>] [-ColumnAlignment <OfficeAlignment[]>] [-AutoFit] [-NoNumericAlignment] [-ShrinkToFit] [-MinimumSize <OfficeLength>] [-KeepTogether] [-KeepWithNext] [-NoBorder] [-NoHeaderBackground] [-NoBanding] [-HeaderRowCount <Int32>] [-RepeatHeaderRowCount <Int32>] [-FooterRowCount <Int32>] [-PassThru] [<CommonParameters>]
```

### PipelineDocument
```powershell
Add-OfficePdfTable [-InputObject] <Object[]> -Document <PdfDocument> [-Property <string[]>] [-Header <string[]>] [-View <OfficeTableView>] [-CollectionSeparator <string>] [-DictionaryEntrySeparator <string>] [-DictionaryKeyValueSeparator <string>] [-MaxCollectionItems <int>] [-MaxNestingDepth <int>] [-ColumnFormat <hashtable>] [-Alignment <OfficeAlignment>] [-TableStyle <string>] [-HeaderBackgroundColor <OfficeColor>] [-HeaderColor <OfficeColor>] [-Color <OfficeColor>] [-BandColor <OfficeColor>] [-BorderColor <OfficeColor>] [-BorderWidth <OfficeLength>] [-Size <OfficeLength>] [-HeaderSize <OfficeLength>] [-LineHeight <Double>] [-CellPaddingX <OfficeLength>] [-CellPaddingY <OfficeLength>] [-SpacingBefore <OfficeLength>] [-SpacingAfter <OfficeLength>] [-Caption <string>] [-CaptionAlignment <OfficeAlignment>] [-CaptionColor <OfficeColor>] [-CaptionSize <OfficeLength>] [-ColumnWidth <OfficeLength[]>] [-ColumnWeight <double[]>] [-ColumnAlignment <OfficeAlignment[]>] [-AutoFit] [-NoNumericAlignment] [-ShrinkToFit] [-MinimumSize <OfficeLength>] [-KeepTogether] [-KeepWithNext] [-NoBorder] [-NoHeaderBackground] [-NoBanding] [-HeaderRowCount <Int32>] [-RepeatHeaderRowCount <Int32>] [-FooterRowCount <Int32>] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Adds a table to a PDF document.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $services = @(
    [pscustomobject]@{ Name = 'Directory'; Status = 'Healthy'; Incidents = 0 }
    [pscustomobject]@{ Name = 'Mail'; Status = 'Watch'; Incidents = 2 }
)
New-OfficePdf -Path .\Examples\Documents\PdfTable.pdf {
    Add-OfficePdfHeading -Text 'Service status'
    Add-OfficePdfTable -InputObject $services -Property Name,Status,Incidents -Header 'Service','Status','Incidents'
}
```

Converts PowerShell objects into a table using selected properties and friendly headers.

## PARAMETERS

### -Alignment
Table alignment.

```yaml
Type: OfficeAlignment
Parameter Sets: Context, Document, PipelineDocument
Aliases: None
Possible values: Left, Center, Right, Justify, Start, End

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -AutoFit
Measure flexible columns from content.

```yaml
Type: SwitchParameter
Parameter Sets: Context, Document, PipelineDocument
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -BandColor
Alternating body row fill color. Named colors and hexadecimal values are accepted.

```yaml
Type: OfficeColor
Parameter Sets: Context, Document, PipelineDocument
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -BorderColor
Border color. Named colors and hexadecimal values are accepted.

```yaml
Type: OfficeColor
Parameter Sets: Context, Document, PipelineDocument
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -BorderWidth
Border width in PDF points.

```yaml
Type: OfficeLength
Parameter Sets: Context, Document, PipelineDocument
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Caption
Caption rendered above the table grid.

```yaml
Type: String
Parameter Sets: Context, Document, PipelineDocument
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -CaptionAlignment
Caption alignment.

```yaml
Type: OfficeAlignment
Parameter Sets: Context, Document, PipelineDocument
Aliases: None
Possible values: Left, Center, Right, Justify, Start, End

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -CaptionColor
Caption color. Named colors and hexadecimal values are accepted.

```yaml
Type: OfficeColor
Parameter Sets: Context, Document, PipelineDocument
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -CaptionSize
Caption font size in PDF points.

```yaml
Type: OfficeLength
Parameter Sets: Context, Document, PipelineDocument
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -CellPaddingX
Horizontal cell padding in PDF points.

```yaml
Type: OfficeLength
Parameter Sets: Context, Document, PipelineDocument
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -CellPaddingY
Vertical cell padding in PDF points.

```yaml
Type: OfficeLength
Parameter Sets: Context, Document, PipelineDocument
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -CollectionSeparator
Text used between items when a property or cell contains a collection.

```yaml
Type: String
Parameter Sets: Context, Document, PipelineDocument
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Color
Body text color. Named colors and hexadecimal values are accepted.

```yaml
Type: OfficeColor
Parameter Sets: Context, Document, PipelineDocument
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ColumnAlignment
Per-column horizontal alignment.

```yaml
Type: OfficeAlignment[]
Parameter Sets: Context, Document, PipelineDocument
Aliases: None
Possible values: Left, Center, Right, Justify, Start, End

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ColumnFormat
Invariant .NET display formats keyed by selected property name.

```yaml
Type: Hashtable
Parameter Sets: Context, Document, PipelineDocument
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ColumnWeight
Relative column width weights.

```yaml
Type: Double[]
Parameter Sets: Context, Document, PipelineDocument
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ColumnWidth
Fixed column widths in PDF points.

```yaml
Type: OfficeLength[]
Parameter Sets: Context, Document, PipelineDocument
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -DictionaryEntrySeparator
Text used between entries when a cell contains a dictionary.

```yaml
Type: String
Parameter Sets: Context, Document, PipelineDocument
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -DictionaryKeyValueSeparator
Text used between a dictionary key and value.

```yaml
Type: String
Parameter Sets: Context, Document, PipelineDocument
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Document
PDF document to update outside the DSL context.

```yaml
Type: PdfDocument
Parameter Sets: Document, PipelineDocument
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -FooterRowCount
Number of trailing rows rendered as footer rows.

```yaml
Type: Int32
Parameter Sets: Context, Document, PipelineDocument
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Header
Header labels. Defaults to property names.

```yaml
Type: String[]
Parameter Sets: Context, Document, PipelineDocument
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -HeaderBackgroundColor
Header fill color. Named colors and hexadecimal values are accepted.

```yaml
Type: OfficeColor
Parameter Sets: Context, Document, PipelineDocument
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -HeaderColor
Header text color. Named colors and hexadecimal values are accepted.

```yaml
Type: OfficeColor
Parameter Sets: Context, Document, PipelineDocument
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -HeaderRowCount
Number of leading rows rendered as header rows.

```yaml
Type: Int32
Parameter Sets: Context, Document, PipelineDocument
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -HeaderSize
Header cell font size in PDF points.

```yaml
Type: OfficeLength
Parameter Sets: Context, Document, PipelineDocument
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -InputObject
Objects or row arrays to render as a table.

```yaml
Type: Object[]
Parameter Sets: Context, Document, PipelineDocument
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -KeepTogether
Keep the table together when possible.

```yaml
Type: SwitchParameter
Parameter Sets: Context, Document, PipelineDocument
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -KeepWithNext
Keep the table with the next block when possible.

```yaml
Type: SwitchParameter
Parameter Sets: Context, Document, PipelineDocument
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -LineHeight
Wrapped line height multiplier for table cells.

```yaml
Type: Double
Parameter Sets: Context, Document, PipelineDocument
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -MaxCollectionItems
Maximum number of items allowed in one nested collection or dictionary cell. Defaults to 1,048,575; increase explicitly for trusted larger values.

```yaml
Type: Int32
Parameter Sets: Context, Document, PipelineDocument
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -MaxNestingDepth
Maximum nesting depth allowed while normalizing one cell value. Defaults to 64; increase explicitly for trusted deeper values.

```yaml
Type: Int32
Parameter Sets: Context, Document, PipelineDocument
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -MinimumSize
Smallest font size, in points, used by -ShrinkToFit.

```yaml
Type: OfficeLength
Parameter Sets: Context, Document, PipelineDocument
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NoBanding
Disable alternating row fill.

```yaml
Type: SwitchParameter
Parameter Sets: Context, Document, PipelineDocument
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NoBorder
Hide table borders.

```yaml
Type: SwitchParameter
Parameter Sets: Context, Document, PipelineDocument
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NoHeaderBackground
Disable the header fill.

```yaml
Type: SwitchParameter
Parameter Sets: Context, Document, PipelineDocument
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NoNumericAlignment
Right-align numeric-looking cell values.

```yaml
Type: SwitchParameter
Parameter Sets: Context, Document, PipelineDocument
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PassThru
Emit the updated document.

```yaml
Type: SwitchParameter
Parameter Sets: Document, PipelineDocument
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Property
Specific object properties to include.

```yaml
Type: String[]
Parameter Sets: Context, Document, PipelineDocument
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -RepeatHeaderRowCount
Number of leading header rows repeated on following pages.

```yaml
Type: Int32
Parameter Sets: Context, Document, PipelineDocument
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ShrinkToFit
Reduce table text size when needed so cell text fits within the resolved cell width.

```yaml
Type: SwitchParameter
Parameter Sets: Context, Document, PipelineDocument
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Size
Body cell font size in PDF points.

```yaml
Type: OfficeLength
Parameter Sets: Context, Document, PipelineDocument
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -SpacingAfter
Spacing after the table in PDF points.

```yaml
Type: OfficeLength
Parameter Sets: Context, Document, PipelineDocument
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -SpacingBefore
Spacing before the table in PDF points.

```yaml
Type: OfficeLength
Parameter Sets: Context, Document, PipelineDocument
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -TableStyle
OfficeIMO table style preset or supported Word table style name.

```yaml
Type: String
Parameter Sets: Context, Document, PipelineDocument
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -View
Projection to apply before writing the table.

```yaml
Type: OfficeTableView
Parameter Sets: Context, Document, PipelineDocument
Aliases: None
Possible values: Normal, Transpose

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `OfficeIMO.Pdf.PdfDocument`
- `System.Object[]`

## OUTPUTS

- `OfficeIMO.Pdf.PdfDocument`

## RELATED LINKS

- None
