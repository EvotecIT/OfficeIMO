---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Add-OfficePdfTable
## SYNOPSIS
Adds a table to a PDF document.

## SYNTAX
### Context (Default)
```powershell
Add-OfficePdfTable [-InputObject] <Object> [-Property <string[]>] [-Header <string[]>] [-View <OfficeTableView>] [-CollectionSeparator <string>] [-DictionaryEntrySeparator <string>] [-DictionaryKeyValueSeparator <string>] [-MaxCollectionItems <int>] [-MaxNestingDepth <int>] [-Align <PdfAlign>] [-TableStyle <string>] [-HeaderFill <string>] [-HeaderTextColor <string>] [-TextColor <string>] [-RowStripeFill <string>] [-BorderColor <string>] [-BorderWidth <Double>] [-FontSize <Double>] [-HeaderFontSize <Double>] [-LineHeight <Double>] [-CellPaddingX <Double>] [-CellPaddingY <Double>] [-SpacingBefore <Double>] [-SpacingAfter <Double>] [-Caption <string>] [-CaptionAlign <PdfAlign>] [-CaptionColor <string>] [-CaptionFontSize <Double>] [-ColumnWidthPoints <double[]>] [-ColumnWidthWeights <double[]>] [-ColumnAlign <PdfColumnAlign[]>] [-AutoFitColumns] [-RightAlignNumeric] [-ShrinkTextToFit] [-MinimumShrinkFontSize <Double>] [-KeepTogether] [-KeepWithNext] [-NoBorder] [-NoHeaderFill] [-NoRowStripeFill] [-HeaderRowCount <Int32>] [-RepeatHeaderRowCount <Int32>] [-FooterRowCount <Int32>] [-PassThru] [<CommonParameters>]
```

### Document
```powershell
Add-OfficePdfTable [-InputObject] <Object> -Document <PdfDocument> [-Property <string[]>] [-Header <string[]>] [-View <OfficeTableView>] [-CollectionSeparator <string>] [-DictionaryEntrySeparator <string>] [-DictionaryKeyValueSeparator <string>] [-MaxCollectionItems <int>] [-MaxNestingDepth <int>] [-Align <PdfAlign>] [-TableStyle <string>] [-HeaderFill <string>] [-HeaderTextColor <string>] [-TextColor <string>] [-RowStripeFill <string>] [-BorderColor <string>] [-BorderWidth <Double>] [-FontSize <Double>] [-HeaderFontSize <Double>] [-LineHeight <Double>] [-CellPaddingX <Double>] [-CellPaddingY <Double>] [-SpacingBefore <Double>] [-SpacingAfter <Double>] [-Caption <string>] [-CaptionAlign <PdfAlign>] [-CaptionColor <string>] [-CaptionFontSize <Double>] [-ColumnWidthPoints <double[]>] [-ColumnWidthWeights <double[]>] [-ColumnAlign <PdfColumnAlign[]>] [-AutoFitColumns] [-RightAlignNumeric] [-ShrinkTextToFit] [-MinimumShrinkFontSize <Double>] [-KeepTogether] [-KeepWithNext] [-NoBorder] [-NoHeaderFill] [-NoRowStripeFill] [-HeaderRowCount <Int32>] [-RepeatHeaderRowCount <Int32>] [-FooterRowCount <Int32>] [-PassThru] [<CommonParameters>]
```

### PipelineDocument
```powershell
Add-OfficePdfTable [-InputObject] <Object> -Document <PdfDocument> [-Property <string[]>] [-Header <string[]>] [-View <OfficeTableView>] [-CollectionSeparator <string>] [-DictionaryEntrySeparator <string>] [-DictionaryKeyValueSeparator <string>] [-MaxCollectionItems <int>] [-MaxNestingDepth <int>] [-Align <PdfAlign>] [-TableStyle <string>] [-HeaderFill <string>] [-HeaderTextColor <string>] [-TextColor <string>] [-RowStripeFill <string>] [-BorderColor <string>] [-BorderWidth <Double>] [-FontSize <Double>] [-HeaderFontSize <Double>] [-LineHeight <Double>] [-CellPaddingX <Double>] [-CellPaddingY <Double>] [-SpacingBefore <Double>] [-SpacingAfter <Double>] [-Caption <string>] [-CaptionAlign <PdfAlign>] [-CaptionColor <string>] [-CaptionFontSize <Double>] [-ColumnWidthPoints <double[]>] [-ColumnWidthWeights <double[]>] [-ColumnAlign <PdfColumnAlign[]>] [-AutoFitColumns] [-RightAlignNumeric] [-ShrinkTextToFit] [-MinimumShrinkFontSize <Double>] [-KeepTogether] [-KeepWithNext] [-NoBorder] [-NoHeaderFill] [-NoRowStripeFill] [-HeaderRowCount <Int32>] [-RepeatHeaderRowCount <Int32>] [-FooterRowCount <Int32>] [-PassThru] [<CommonParameters>]
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

### -Align
Table alignment.

```yaml
Type: PdfAlign
Parameter Sets: Context, Document, PipelineDocument
Aliases: None
Possible values: Left, Center, Right, Justify

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -AutoFitColumns
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

### -BorderColor
Border color. Named colors and hexadecimal values are accepted.

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

### -BorderWidth
Border width in PDF points.

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

### -CaptionAlign
Caption alignment.

```yaml
Type: PdfAlign
Parameter Sets: Context, Document, PipelineDocument
Aliases: None
Possible values: Left, Center, Right, Justify

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -CaptionColor
Caption color. Named colors and hexadecimal values are accepted.

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

### -CaptionFontSize
Caption font size in PDF points.

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

### -CellPaddingX
Horizontal cell padding in PDF points.

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

### -CellPaddingY
Vertical cell padding in PDF points.

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

### -ColumnAlign
Per-column horizontal alignment.

```yaml
Type: PdfColumnAlign[]
Parameter Sets: Context, Document, PipelineDocument
Aliases: None
Possible values: Left, Center, Right

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ColumnWidthPoints
Fixed column widths in PDF points.

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

### -ColumnWidthWeights
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

### -FontSize
Body cell font size in PDF points.

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

### -HeaderFill
Header fill color. Named colors and hexadecimal values are accepted.

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

### -HeaderFontSize
Header cell font size in PDF points.

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

### -HeaderTextColor
Header text color. Named colors and hexadecimal values are accepted.

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

### -InputObject
Objects or row arrays to render as a table.

```yaml
Type: Object
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

### -MinimumShrinkFontSize
Smallest font size, in points, used by -ShrinkTextToFit.

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

### -NoHeaderFill
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

### -NoRowStripeFill
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

### -PassThru
Emit the updated document.

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

### -RightAlignNumeric
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

### -RowStripeFill
Alternating body row fill color. Named colors and hexadecimal values are accepted.

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

### -ShrinkTextToFit
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

### -SpacingAfter
Spacing after the table in PDF points.

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

### -SpacingBefore
Spacing before the table in PDF points.

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

### -TextColor
Body text color. Named colors and hexadecimal values are accepted.

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
- `System.Object`

## OUTPUTS

- `OfficeIMO.Pdf.PdfDocument`

## RELATED LINKS

- None
