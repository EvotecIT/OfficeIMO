---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# New-OfficeConversionOptions
## SYNOPSIS
Creates discoverable renderer and reconstruction settings for general conversion jobs.

## SYNTAX
### __AllParameterSets
```powershell
New-OfficeConversionOptions [-WordOptions <WordToPdfOptions>] [-ExcelOptions <ExcelToPdfOptions>] [-PowerPointOptions <PowerPointToPdfOptions>] [-HtmlOptions <HtmlToPdfOptions>] [-MarkdownOptions <MarkdownToPdfOptions>] [-RtfOptions <RtfToPdfOptions>] [-TextOptions <PdfPlainTextOptions>] [-DrawOptions <OdgToPdfOptions>] [-VisioOptions <VisioToPdfOptions>] [-XpsOptions <XpsToPdfOptions>] [-ChmOptions <ChmConversionOptions>] [-DjVuOptions <DjVuToPdfOptions>] [-PublisherReadOptions <PublisherReadOptions>] [-PublisherPdfOptions <PdfOptions>] [-CompressPdfOutput] [-WorksheetLayout <ExcelPdfWorksheetLayoutMode>] [-PdfHtmlProfile <PdfHtmlProfile>] [-RequireNoLoss] [-AllowLegacyImportLoss] [-PageRange <string>] [-PdfWordMode <PdfWordImportMode>] [-PdfPowerPointMode <PdfPowerPointImportMode>] [-RasterDpi <Double>] [<CommonParameters>]
```

## DESCRIPTION
Requires PowerShell 7.4 or newer. Format-specific options apply only to their matching routes.

## EXAMPLES

### EXAMPLE 1
```powershell
$options = New-OfficeConversionOptions -RequireNoLoss -ExcelOptions (New-OfficeExcelPdfOptions -UseWorksheetCharts)
ConvertTo-OfficeDocument -InputPath ./report.xlsx -OutputPath ./report.pdf -ConversionOptions $options
```


## PARAMETERS

### -AllowLegacyImportLoss
Explicitly accept known legacy DOC import reductions.

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

### -ChmOptions
CHM topic selection and conversion budgets.

```yaml
Type: ChmConversionOptions
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -CompressPdfOutput
Losslessly compress generated PDF output.

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

### -DjVuOptions
DjVu scanned-page PDF and stored-text settings.

```yaml
Type: DjVuToPdfOptions
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -DrawOptions
ODG/FODG drawing projection and PDF options; accepts a property hashtable.

```yaml
Type: OdgToPdfOptions
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ExcelOptions
Excel-to-PDF settings from New-OfficeExcelPdfOptions.

```yaml
Type: ExcelToPdfOptions
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -HtmlOptions
HTML-to-PDF settings from New-OfficeHtmlPdfOptions.

```yaml
Type: HtmlToPdfOptions
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -MarkdownOptions
Markdown-to-PDF settings from New-OfficeMarkdownPdfOptions.

```yaml
Type: MarkdownToPdfOptions
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PageRange
Page ranges for PDF source reconstruction.

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

### -PdfHtmlProfile
Semantic or positioned-review HTML reconstructed from PDF.

```yaml
Type: PdfHtmlProfile
Parameter Sets: __AllParameterSets
Aliases: None
Possible values: Semantic, PositionedReview

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PdfPowerPointMode
PDF-to-PowerPoint reconstruction or rendered-page strategy.

```yaml
Type: PdfPowerPointImportMode
Parameter Sets: __AllParameterSets
Aliases: None
Possible values: VisualPages, EditableTables, HybridVisualAndEditableTables, EditableContent, Auto

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PdfWordMode
PDF-to-Word reconstruction or rendered-page strategy.

```yaml
Type: PdfWordImportMode
Parameter Sets: __AllParameterSets
Aliases: None
Possible values: EditableContent, VisualPages

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PowerPointOptions
PowerPoint-to-PDF settings from New-OfficePowerPointPdfOptions.

```yaml
Type: PowerPointToPdfOptions
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PublisherPdfOptions
Shared PDF settings for recovered Publisher pages.

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

### -PublisherReadOptions
Native Publisher decoding and projection limits.

```yaml
Type: PublisherReadOptions
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -RasterDpi
Resolution of rendered PDF pages, when applicable.

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

### -RequireNoLoss
Reject any reduction reported by the selected route.

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

### -RtfOptions
RTF-to-PDF settings from New-OfficeRtfPdfOptions.

```yaml
Type: RtfToPdfOptions
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -TextOptions
Literal-text settings from New-OfficeTextPdfOptions.

```yaml
Type: PdfPlainTextOptions
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -VisioOptions
Visio diagram-page projection settings; accepts a property hashtable.

```yaml
Type: VisioToPdfOptions
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -WordOptions
Word-to-PDF settings from New-OfficeWordPdfOptions.

```yaml
Type: WordToPdfOptions
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
Workbook worksheet-canvas or flowing-table layout.

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

### -XpsOptions
Native XPS preservation and PDF output settings.

```yaml
Type: XpsToPdfOptions
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

- `None`

## OUTPUTS

- `None`

## RELATED LINKS

- None
