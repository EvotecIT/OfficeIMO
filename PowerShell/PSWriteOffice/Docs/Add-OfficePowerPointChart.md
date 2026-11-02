---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Add-OfficePowerPointChart
## SYNOPSIS
Adds a chart to a PowerPoint slide.

## SYNTAX
### Sample (Default)
```powershell
Add-OfficePowerPointChart [-Slide <PowerPointSlideSelector>] [-Type <OfficeChartType>] [-X <OfficeLength>] [-Y <OfficeLength>] [-Width <OfficeLength>] [-Height <OfficeLength>] [-Column <Int32>] [-Span <Int32>] [-Title <string>] [-PassThru] [<CommonParameters>]
```

### InputObject
```powershell
Add-OfficePowerPointChart -InputObject <Object[]> -Category <string> -Value <string[]> [-Slide <PowerPointSlideSelector>] [-Type <OfficeChartType>] [-X <OfficeLength>] [-Y <OfficeLength>] [-Width <OfficeLength>] [-Height <OfficeLength>] [-Column <Int32>] [-Span <Int32>] [-Title <string>] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Supports default chart data or object-based category/series mappings for standard and scatter charts.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $rows = @(
    [pscustomobject]@{ Month = 'Jan'; Sales = 42; Profit = 9 }
    [pscustomobject]@{ Month = 'Feb'; Sales = 55; Profit = 13 }
)
New-OfficePowerPoint -Path .\Examples\Documents\PowerPointChart.pptx {
    $slide = Add-OfficePowerPointSlide -PassThru
    Add-OfficePowerPointChart -Slide $slide -InputObject $rows -Category Month -Value Sales,Profit -Title 'Monthly performance'
}
```

Creates a clustered column chart using Month for categories and Sales/Profit as series.

### EXAMPLE 2
```powershell
PS> $rows = @(
    [pscustomobject]@{ Quarter = 1; Revenue = 20 }
    [pscustomobject]@{ Quarter = 2; Revenue = 34 }
)
New-OfficePowerPoint -Path .\Examples\Documents\PowerPointScatter.pptx {
    $slide = Add-OfficePowerPointSlide -PassThru
    Add-OfficePowerPointChart -Slide $slide -Type Scatter -InputObject $rows -Category Quarter -Value Revenue -Title 'Revenue trend'
}
```

Creates a scatter chart using Quarter on the X axis and Revenue on the Y axis.

## PARAMETERS

### -Category
Property used for category labels, or numeric X values on scatter charts.

```yaml
Type: String
Parameter Sets: InputObject
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Column
One-based start column in the native twelve-column grid; explicit X overrides it.

```yaml
Type: Int32
Parameter Sets: Sample, InputObject
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Height
Chart height; numbers mean points.

```yaml
Type: OfficeLength
Parameter Sets: Sample, InputObject
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
Parameter Sets: InputObject
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -PassThru
Emit the object created or changed by the command.

```yaml
Type: SwitchParameter
Parameter Sets: Sample, InputObject
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Slide
Target slide that will receive the chart (optional inside DSL).

```yaml
Type: PowerPointSlideSelector
Parameter Sets: Sample, InputObject
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Span
Number of grid columns; defaults to twelve, and explicit Width overrides it.

```yaml
Type: Int32
Parameter Sets: Sample, InputObject
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Title
Optional chart title.

```yaml
Type: String
Parameter Sets: Sample, InputObject
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
Type: OfficeChartType
Parameter Sets: Sample, InputObject
Aliases: None
Possible values: ColumnClustered, Column, ColumnStacked, ColumnStacked100, BarClustered, Bar, BarStacked, BarStacked100, Line, LineStacked, LineStacked100, Area, AreaStacked, AreaStacked100, Scatter, Radar, Pie, Doughnut, Bubble, Column3DClustered, Column3DStacked, Column3DStacked100, Bar3DClustered, Bar3DStacked, Bar3DStacked100, Line3D, Area3D, Area3DStacked, Area3DStacked100, Pie3D, PieOfPie, BarOfPie, Stock, Surface, SurfaceWireframe, SurfaceContour, Funnel, Waterfall, BoxWhisker, Treemap, SurfaceContourWireframe, Sunburst

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Value
Properties used as numeric series, or Y series on scatter charts.

```yaml
Type: String[]
Parameter Sets: InputObject
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Width
Chart width; numbers mean points.

```yaml
Type: OfficeLength
Parameter Sets: Sample, InputObject
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -X
Left offset from the slide origin; numbers mean points.

```yaml
Type: OfficeLength
Parameter Sets: Sample, InputObject
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Y
Top offset from the slide origin; numbers mean points.

```yaml
Type: OfficeLength
Parameter Sets: Sample, InputObject
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

- `System.Object[]`

## OUTPUTS

- `OfficeIMO.PowerPoint.PowerPointChart`

## RELATED LINKS

- None
