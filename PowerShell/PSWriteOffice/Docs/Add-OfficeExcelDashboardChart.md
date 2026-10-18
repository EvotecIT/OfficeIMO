---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Add-OfficeExcelDashboardChart
## SYNOPSIS
Adds a dashboard-ready chart using an OfficeIMO chart preset.

## SYNTAX
### ContextRange (Default)
```powershell
Add-OfficeExcelDashboardChart [-Range] <string> -Row <int> -Column <int> [-Preset <ExcelDashboardChartPreset>] [-ChartType <ExcelChartType>] [-Title <string>] [-HasHeaders <bool>] [-IncludeCachedData <bool>] [-WidthPixels <Int32>] [-HeightPixels <Int32>] [-StyleId <Int32>] [-ColorStyleId <Int32>] [-PassThru] [<CommonParameters>]
```

### DocumentRange
```powershell
Add-OfficeExcelDashboardChart [-Range] <string> -Document <ExcelDocument> -Row <int> -Column <int> [-Sheet <string>] [-SheetIndex <Int32>] [-Preset <ExcelDashboardChartPreset>] [-ChartType <ExcelChartType>] [-Title <string>] [-HasHeaders <bool>] [-IncludeCachedData <bool>] [-WidthPixels <Int32>] [-HeightPixels <Int32>] [-StyleId <Int32>] [-ColorStyleId <Int32>] [-PassThru] [<CommonParameters>]
```

### DocumentTable
```powershell
Add-OfficeExcelDashboardChart [-TableName] <string> -Document <ExcelDocument> -Row <int> -Column <int> [-Sheet <string>] [-SheetIndex <Int32>] [-Preset <ExcelDashboardChartPreset>] [-ChartType <ExcelChartType>] [-Title <string>] [-IncludeCachedData <bool>] [-WidthPixels <Int32>] [-HeightPixels <Int32>] [-StyleId <Int32>] [-ColorStyleId <Int32>] [-PassThru] [<CommonParameters>]
```

### ContextTable
```powershell
Add-OfficeExcelDashboardChart [-TableName] <string> -Row <int> -Column <int> [-Preset <ExcelDashboardChartPreset>] [-ChartType <ExcelChartType>] [-Title <string>] [-IncludeCachedData <bool>] [-WidthPixels <Int32>] [-HeightPixels <Int32>] [-StyleId <Int32>] [-ColorStyleId <Int32>] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Adds a dashboard-ready chart using an OfficeIMO chart preset.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> ExcelSheet 'Dashboard' { Add-OfficeExcelDashboardChart -Range A1:B12 -Preset CompactComparison -Row 1 -Column 5 -Title 'Revenue' }
```

Creates a styled chart from the range using reusable OfficeIMO dashboard chart defaults.

## PARAMETERS

### -ChartType
Optional chart type override.

```yaml
Type: ExcelChartType
Parameter Sets: ContextRange, DocumentRange, DocumentTable, ContextTable
Aliases: None
Possible values: ColumnClustered, ColumnStacked, ColumnStacked100, Column3DClustered, Column3DStacked, Column3DStacked100, BarClustered, BarStacked, BarStacked100, Bar3DClustered, Bar3DStacked, Bar3DStacked100, Line, LineStacked, LineStacked100, Line3D, Area, AreaStacked, AreaStacked100, Area3D, Area3DStacked, Area3DStacked100, Pie, Pie3D, PieOfPie, BarOfPie, Doughnut, Scatter, Bubble, Radar, Stock, Surface, SurfaceWireframe, SurfaceContour, SurfaceContourWireframe

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ColorStyleId
Optional chart color style id override.

```yaml
Type: Int32
Parameter Sets: ContextRange, DocumentRange, DocumentTable, ContextTable
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Column
Top-left column (1-based) where the chart should be placed.

```yaml
Type: Int32
Parameter Sets: ContextRange, DocumentRange, DocumentTable, ContextTable
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Document
Workbook to operate on outside the DSL context.

```yaml
Type: ExcelDocument
Parameter Sets: DocumentRange, DocumentTable
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -HasHeaders
Whether the range includes headers.

```yaml
Type: Boolean
Parameter Sets: ContextRange, DocumentRange
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -HeightPixels
Optional chart height in pixels.

```yaml
Type: Int32
Parameter Sets: ContextRange, DocumentRange, DocumentTable, ContextTable
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -IncludeCachedData
Include cached data in the chart for portability.

```yaml
Type: Boolean
Parameter Sets: ContextRange, DocumentRange, DocumentTable, ContextTable
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PassThru
Emit the created chart.

```yaml
Type: SwitchParameter
Parameter Sets: ContextRange, DocumentRange, DocumentTable, ContextTable
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Preset
Dashboard chart preset.

```yaml
Type: ExcelDashboardChartPreset
Parameter Sets: ContextRange, DocumentRange, DocumentTable, ContextTable
Aliases: None
Possible values: Comparison, Trend, Contribution, CompactComparison

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Range
A1 range containing chart data.

```yaml
Type: String
Parameter Sets: ContextRange, DocumentRange
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Row
Top-left row (1-based) where the chart should be placed.

```yaml
Type: Int32
Parameter Sets: ContextRange, DocumentRange, DocumentTable, ContextTable
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Sheet
Worksheet name when using Document.

```yaml
Type: String
Parameter Sets: DocumentRange, DocumentTable
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -SheetIndex
Worksheet index (0-based) when using Document.

```yaml
Type: Int32
Parameter Sets: DocumentRange, DocumentTable
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -StyleId
Optional chart style id override.

```yaml
Type: Int32
Parameter Sets: ContextRange, DocumentRange, DocumentTable, ContextTable
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -TableName
Table name containing chart data.

```yaml
Type: String
Parameter Sets: DocumentTable, ContextTable
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Title
Chart title.

```yaml
Type: String
Parameter Sets: ContextRange, DocumentRange, DocumentTable, ContextTable
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -WidthPixels
Optional chart width in pixels.

```yaml
Type: Int32
Parameter Sets: ContextRange, DocumentRange, DocumentTable, ContextTable
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

- `OfficeIMO.Excel.ExcelDocument`

## OUTPUTS

- `None`

## RELATED LINKS

- None
