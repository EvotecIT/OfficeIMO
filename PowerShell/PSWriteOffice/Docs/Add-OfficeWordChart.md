---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Add-OfficeWordChart
## SYNOPSIS
Adds a chart to a Word document.

## SYNTAX
### Context (Default)
```powershell
Add-OfficeWordChart -InputObject <Object[]> -CategoryProperty <string> -SeriesProperty <string[]> [-Type <WordChartType>] [-WidthPixels <int>] [-HeightPixels <int>] [-Title <string>] [-SeriesColor <string[]>] [-Legend] [-LegendPosition <OfficeChartLegendPosition>] [-XAxisTitle <string>] [-YAxisTitle <string>] [-FitToPageWidth] [-WidthFraction <double>] [-PassThru] [<CommonParameters>]
```

### Document
```powershell
Add-OfficeWordChart -InputObject <Object[]> -CategoryProperty <string> -SeriesProperty <string[]> [-Document <WordDocument>] [-Type <WordChartType>] [-WidthPixels <int>] [-HeightPixels <int>] [-Title <string>] [-SeriesColor <string[]>] [-Legend] [-LegendPosition <OfficeChartLegendPosition>] [-XAxisTitle <string>] [-YAxisTitle <string>] [-FitToPageWidth] [-WidthFraction <double>] [-PassThru] [<CommonParameters>]
```

### Paragraph
```powershell
Add-OfficeWordChart -InputObject <Object[]> -CategoryProperty <string> -SeriesProperty <string[]> [-Paragraph <WordParagraph>] [-Type <WordChartType>] [-WidthPixels <int>] [-HeightPixels <int>] [-Title <string>] [-SeriesColor <string[]>] [-Legend] [-LegendPosition <OfficeChartLegendPosition>] [-XAxisTitle <string>] [-YAxisTitle <string>] [-FitToPageWidth] [-WidthFraction <double>] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Creates a Word chart from object data using one category property and one or more numeric series properties.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $rows = @(
    [pscustomobject]@{ Region = 'North America'; Revenue = 125000; Profit = 42000 }
    [pscustomobject]@{ Region = 'EMEA'; Revenue = 98000; Profit = 31000 }
    [pscustomobject]@{ Region = 'APAC'; Revenue = 143000; Profit = 52000 }
)
New-OfficeWord -Path .\RegionalReport.docx {
    Add-OfficeWordParagraph -Text 'Regional revenue'
    Add-OfficeWordChart -Type Pie -InputObject $rows -CategoryProperty Region -SeriesProperty Revenue -Title 'Revenue mix' -FitToPageWidth -WidthFraction 0.75
    Add-OfficeWordChart -Type Bar -InputObject $rows -CategoryProperty Region -SeriesProperty Revenue, Profit -Legend -XAxisTitle 'Region' -YAxisTitle 'Amount'
}
```

Creates a report with a pie chart and a multi-series bar chart from PowerShell objects.

### EXAMPLE 2
```powershell
PS> $trend = @(
    [pscustomobject]@{ Month = 'Jan'; Sales = 10; Profit = 4 }
    [pscustomobject]@{ Month = 'Feb'; Sales = 12; Profit = 5 }
    [pscustomobject]@{ Month = 'Mar'; Sales = 15; Profit = 7 }
)
$doc = New-OfficeWord -Path .\Trend.docx -NoSave
Add-OfficeWordChart -Document $doc -Type Line -InputObject $trend -CategoryProperty Month -SeriesProperty Sales, Profit -Legend -Title 'Quarter trend'
Save-OfficeWord -Document $doc
```

Creates a multi-series line chart on the document and shows a legend.

## PARAMETERS

### -CategoryProperty
Property name used for category labels.

```yaml
Type: String
Parameter Sets: Context, Document, Paragraph
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Document
Target document that will receive the chart.

```yaml
Type: WordDocument
Parameter Sets: Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -FitToPageWidth
Scale the chart width to the page content width.

```yaml
Type: SwitchParameter
Parameter Sets: Context, Document, Paragraph
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -HeightPixels
Chart height in pixels.

```yaml
Type: Int32
Parameter Sets: Context, Document, Paragraph
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -InputObject
Source objects used to build chart data.

```yaml
Type: Object[]
Parameter Sets: Context, Document, Paragraph
Aliases: Data
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Legend
Add a legend to the chart.

```yaml
Type: SwitchParameter
Parameter Sets: Context, Document, Paragraph
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -LegendPosition
Legend position when -Legend is used.

```yaml
Type: OfficeChartLegendPosition
Parameter Sets: Context, Document, Paragraph
Aliases: None
Possible values: Bottom, TopRight, Left, Right, Top

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Paragraph
Target paragraph used as the chart anchor.

```yaml
Type: WordParagraph
Parameter Sets: Paragraph
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -PassThru
Emit the created chart.

```yaml
Type: SwitchParameter
Parameter Sets: Context, Document, Paragraph
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -SeriesColor
Color values applied to the series in order.

```yaml
Type: String[]
Parameter Sets: Context, Document, Paragraph
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -SeriesProperty
Property names used as numeric series.

```yaml
Type: String[]
Parameter Sets: Context, Document, Paragraph
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Title
Optional chart title.

```yaml
Type: String
Parameter Sets: Context, Document, Paragraph
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Type
Chart type to create.

```yaml
Type: WordChartType
Parameter Sets: Context, Document, Paragraph
Aliases: None
Possible values: Pie, Bar, Line, Area

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -WidthFraction
Fraction of the page content width to use when -FitToPageWidth is specified.

```yaml
Type: Double
Parameter Sets: Context, Document, Paragraph
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -WidthPixels
Chart width in pixels.

```yaml
Type: Int32
Parameter Sets: Context, Document, Paragraph
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -XAxisTitle
Optional X axis title for non-pie charts.

```yaml
Type: String
Parameter Sets: Context, Document, Paragraph
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -YAxisTitle
Optional Y axis title for non-pie charts.

```yaml
Type: String
Parameter Sets: Context, Document, Paragraph
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

- `OfficeIMO.Word.WordDocument`
- `OfficeIMO.Word.WordParagraph`

## OUTPUTS

- `OfficeIMO.Word.WordChart`

## RELATED LINKS

- None
