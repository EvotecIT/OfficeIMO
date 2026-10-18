---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# New-OfficeExcelPdfOptions
## SYNOPSIS
Creates discoverable Excel-to-PDF conversion options for Export-OfficeDocumentPdf.

## SYNTAX
### __AllParameterSets
```powershell
New-OfficeExcelPdfOptions [-PdfOptions <PdfOptions>] [-FontFamily <string>] [-PageSize <PageSize>] [-MarginLeft <Double>] [-MarginTop <Double>] [-MarginRight <Double>] [-MarginBottom <Double>] [-WorksheetLayout <ExcelPdfWorksheetLayoutMode>] [-SheetName <string[]>] [-RespectWorkbookSheetVisibility] [-UseWorksheetPrintAreas] [-UseWorksheetPageSetup] [-UseWorksheetPrintTitleRows] [-UseWorksheetPageBreaks] [-UseWorksheetHeadersAndFooters] [-UseWorksheetHeaderFooterImages] [-UseWorksheetCellStyles] [-UseWorksheetHyperlinks] [-UseWorksheetImages] [-UseWorksheetCharts] [-ChartStyle <OfficeChartStyle>] [-ChartLayout <OfficeChartLayout>] [-UseWorksheetMergedCells] [-UseWorksheetColumnWidths] [-UseWorksheetRowHeights] [-RespectWorksheetHiddenRowsAndColumns] [-IncludeSheetHeadings] [-HeaderRowCount <Int32>] [-MaxRowsPerSheet <Int32>] [-UseBoundedWorksheetRead] [-EmptyCellText <string>] [-AllowSystemFontEmbedding] [-AllowDocumentFontEmbedding] [<CommonParameters>]
```

## DESCRIPTION
Creates discoverable Excel-to-PDF conversion options for Export-OfficeDocumentPdf.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $options = New-OfficeExcelPdfOptions -SheetName Summary,Services -UseWorksheetCharts -UseWorksheetImages
Export-OfficeDocumentPdf -InputPath .\Report.xlsx -Path .\Report.pdf -ExcelOptions $options
```


## PARAMETERS

### -AllowDocumentFontEmbedding
Allow embedding installed fonts selected by a trusted workbook. System-font embedding must also be enabled.

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

### -AllowSystemFontEmbedding
Allow embedding fonts discovered on the current system.

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

### -ChartLayout
Chart layout override.

```yaml
Type: OfficeChartLayout
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ChartStyle
Chart visual style override.

```yaml
Type: OfficeChartStyle
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -EmptyCellText
Text used when a worksheet cell is empty.

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

### -FontFamily
Default font family used when the workbook does not specify one.

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

### -HeaderRowCount
Number of leading rows treated as headers.

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

### -IncludeSheetHeadings
Include worksheet row and column headings.

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

### -MarginBottom
Bottom page margin in PDF points.

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

### -MarginLeft
Left page margin in PDF points.

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

### -MarginRight
Right page margin in PDF points.

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

### -MarginTop
Top page margin in PDF points.

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

### -MaxRowsPerSheet
Maximum worksheet rows to read and render.

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

### -PageSize
PDF page size.

```yaml
Type: PageSize
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PdfOptions
Underlying low-level OfficeIMO PDF options.

```yaml
Type: PdfOptions
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -RespectWorkbookSheetVisibility
Exclude workbook sheets marked hidden.

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

### -RespectWorksheetHiddenRowsAndColumns
Exclude hidden worksheet rows and columns.

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

### -SheetName
Worksheet names to export. The default exports all eligible sheets.

```yaml
Type: String[]
Parameter Sets: __AllParameterSets
Aliases: SheetNames
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -UseBoundedWorksheetRead
Use bounded worksheet reads for large workbooks.

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

### -UseWorksheetCellStyles
Render worksheet cell styles.

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

### -UseWorksheetCharts
Render worksheet charts.

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

### -UseWorksheetColumnWidths
Honor worksheet column widths.

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

### -UseWorksheetHeaderFooterImages
Render images referenced by worksheet headers and footers.

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

### -UseWorksheetHeadersAndFooters
Render worksheet headers and footers.

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

### -UseWorksheetHyperlinks
Render worksheet hyperlinks.

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

### -UseWorksheetImages
Render worksheet images.

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

### -UseWorksheetMergedCells
Render merged worksheet cells.

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

### -UseWorksheetPageBreaks
Honor worksheet page breaks.

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

### -UseWorksheetPageSetup
Honor worksheet page setup.

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

### -UseWorksheetPrintAreas
Honor worksheet print areas.

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

### -UseWorksheetPrintTitleRows
Honor worksheet rows configured to repeat on printed pages.

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

### -UseWorksheetRowHeights
Honor worksheet row heights.

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

### -WorksheetLayout
Controls how worksheet content is laid out on PDF pages.

```yaml
Type: ExcelPdfWorksheetLayoutMode
Parameter Sets: __AllParameterSets
Aliases: None
Possible values: WorksheetCanvas, FlowTable

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

- `OfficeIMO.Excel.Pdf.ExcelToPdfOptions`

## RELATED LINKS

- None
