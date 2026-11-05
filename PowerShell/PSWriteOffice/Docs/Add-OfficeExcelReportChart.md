---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Add-OfficeExcelReportChart
## SYNOPSIS
Adds a chart of a report table, using a dashboard preset or chart type.

## SYNTAX
### Default (Default)
```powershell
Add-OfficeExcelReportChart [-Table <string>] [-Type <OfficeChartType>] [-Preset <ExcelDashboardChartPreset>] [-Title <string>] [-Category <string>] [-Value <string[]>] [-Width <OfficeLength>] [-Height <OfficeLength>] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Adds a chart of a report table, using a dashboard preset or chart type.

## EXAMPLES

### EXAMPLE 1
```powershell
New-OfficeExcel -Path .\Reports\Regions.xlsx {
    ExcelReportSheet 'Summary' {
        ExcelReportTable (Import-OfficeCsv -Path .\Data\Region-Revenue.csv) 'Revenue by region'
        ExcelReportChart -Category Region -Value Revenue
    }
}
```

Charts the previous report table and places the chart beside it.

### EXAMPLE 2
```powershell
New-OfficeExcel -Path .\Reports\Helpdesk-Review.xlsx {
    ExcelReportSheet 'Trend' {
        ExcelReportTitle 'Helpdesk trend' 'Last 12 weeks'
        ExcelReportTable (Import-OfficeCsv -Path .\Data\Ticket-Trend.csv) 'Tickets per week'
        ExcelReportChart -Preset Trend -Category Week -Value Opened, Closed -Title 'Opened and closed tickets'
    }
}
```

Adds a weekly ticket table and charts opened against closed tickets with the trend preset.

## PARAMETERS

### -Category
Category column; defaults to the first table column.

```yaml
Type: String
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
Chart height; defaults to 360 pixels. Bare numbers mean points.

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

### -PassThru
Emits the native result after the operation.

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

### -Preset
Dashboard appearance preset from the native Excel chart owner.

```yaml
Type: ExcelDashboardChartPreset
Parameter Sets: Default
Aliases: None
Possible values: Comparison, Trend, Contribution, CompactComparison

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Table
Native Excel table used to resolve the source range or header.

```yaml
Type: String
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Title
Chart title.

```yaml
Type: String
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Type
Shared chart family. Bubble requires the separate native bubble-data contract.

```yaml
Type: OfficeChartType
Parameter Sets: Default
Aliases: None
Possible values: ColumnClustered, Column, ColumnStacked, ColumnStacked100, BarClustered, Bar, BarStacked, BarStacked100, Line, LineStacked, LineStacked100, Area, AreaStacked, AreaStacked100, Scatter, Radar, Pie, Doughnut, Bubble, Column3DClustered, Column3DStacked, Column3DStacked100, Bar3DClustered, Bar3DStacked, Bar3DStacked100, Line3D, Area3D, Area3DStacked, Area3DStacked100, Pie3D, PieOfPie, BarOfPie, Stock, Surface, SurfaceWireframe, SurfaceContour, Funnel, Waterfall, BoxWhisker, Treemap, SurfaceContourWireframe, Sunburst

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Value
Value or values consumed by the selected operation.

```yaml
Type: String[]
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Width
Chart width; defaults to 640 pixels. Bare numbers mean points.

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

- `None`

## RELATED LINKS

- None
