---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Add-OfficeExcelChart
## SYNOPSIS
Adds a native Excel chart from a range, named table or input rows.

## SYNTAX
### Range (Default)
```powershell
Add-OfficeExcelChart [-Range] <string> -Address <string> [-Document <ExcelDocument>] [-Sheet <OfficeSheetSelector>] [-Width <OfficeLength>] [-Height <OfficeLength>] [-Type <OfficeChartType>] [-Title <string>] [-Preset <ExcelDashboardChartPreset>] [-Style <Int32>] [-ColorStyle <Int32>] [-NoHeader] [-NoCachedData] [-PassThru] [<CommonParameters>]
```

### Table
```powershell
Add-OfficeExcelChart -Table <string> -Address <string> [-Document <ExcelDocument>] [-Sheet <OfficeSheetSelector>] [-Category <string>] [-Value <string[]>] [-Width <OfficeLength>] [-Height <OfficeLength>] [-Type <OfficeChartType>] [-Title <string>] [-Preset <ExcelDashboardChartPreset>] [-Style <Int32>] [-ColorStyle <Int32>] [-NoCachedData] [-PassThru] [<CommonParameters>]
```

### Data
```powershell
Add-OfficeExcelChart -InputObject <Object> -Category <string> -Value <string[]> -Address <string> [-Document <ExcelDocument>] [-Sheet <OfficeSheetSelector>] [-Width <OfficeLength>] [-Height <OfficeLength>] [-Type <OfficeChartType>] [-Title <string>] [-Preset <ExcelDashboardChartPreset>] [-Style <Int32>] [-ColorStyle <Int32>] [-NoCachedData] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Table columns are selected by header. With no value selection, every numeric column except the category is plotted. Input rows are written as a table below existing worksheet data. Output is quiet unless PassThru is specified.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> ExcelSheet 'Sales' { ExcelChart -Table Sales -Category Region -Value Revenue,Profit -Address F2 -Type ColumnClustered }
```

Places a chart at F2 using only the named table's data rows.

### EXAMPLE 2
```powershell
PS> $rows | Add-OfficeExcelChart -Document $book -Sheet Sales -Category Region -Value Revenue -Address F2 -Width 18cm -Height 9cm
```

Writes a table on Sales and adds a native chart with length-based dimensions.

## PARAMETERS

### -Address
Single A1 cell anchoring the top-left corner of the chart.

```yaml
Type: String
Parameter Sets: Range, Table, Data
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Category
Category column; defaults to the first table column.

```yaml
Type: String
Parameter Sets: Table, Data
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ColorStyle
Built-in chart color-style identifier.

```yaml
Type: Int32
Parameter Sets: Range, Table, Data
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Document
Workbook to modify outside an Excel composition block.

```yaml
Type: ExcelDocument
Parameter Sets: Range, Table, Data
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -Height
Chart height; defaults to 360 pixels. Bare numbers mean points.

```yaml
Type: OfficeLength
Parameter Sets: Range, Table, Data
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -InputObject
Objects or tabular input to write as a table and chart.

```yaml
Type: Object
Parameter Sets: Data
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -NoCachedData
Omit cached chart values. Requires a contiguous category/value source block.

```yaml
Type: SwitchParameter
Parameter Sets: Range, Table, Data
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NoHeader
Treat the first row of Range as data rather than column headers.

```yaml
Type: SwitchParameter
Parameter Sets: Range
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PassThru
Return the native ExcelChart.

```yaml
Type: SwitchParameter
Parameter Sets: Range, Table, Data
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
Parameter Sets: Range, Table, Data
Aliases: None
Possible values: Comparison, Trend, Contribution, CompactComparison

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Range
Bounded A1 range containing categories followed by value columns.

```yaml
Type: String
Parameter Sets: Range
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Sheet
Sheet name, one-based position or an existing sheet belonging to Document.

```yaml
Type: OfficeSheetSelector
Parameter Sets: Range, Table, Data
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Style
Built-in chart style identifier.

```yaml
Type: Int32
Parameter Sets: Range, Table, Data
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Table
Name or display name of a table on the selected worksheet.

```yaml
Type: String
Parameter Sets: Table
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Title
Chart title.

```yaml
Type: String
Parameter Sets: Range, Table, Data
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
Parameter Sets: Range, Table, Data
Aliases: None
Possible values: ColumnClustered, Column, ColumnStacked, ColumnStacked100, BarClustered, Bar, BarStacked, BarStacked100, Line, LineStacked, LineStacked100, Area, AreaStacked, AreaStacked100, Scatter, Radar, Pie, Doughnut, Bubble, Column3DClustered, Column3DStacked, Column3DStacked100, Bar3DClustered, Bar3DStacked, Bar3DStacked100, Line3D, Area3D, Area3DStacked, Area3DStacked100, Pie3D, PieOfPie, BarOfPie, Stock, Surface, SurfaceWireframe, SurfaceContour, Funnel, Waterfall, BoxWhisker, Treemap, SurfaceContourWireframe, Sunburst

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Value
Value columns in series order; defaults to all numeric table columns except Category.

```yaml
Type: String[]
Parameter Sets: Table, Data
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
Parameter Sets: Range, Table, Data
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
- `System.Object`

## OUTPUTS

- `OfficeIMO.Excel.ExcelChart`

## RELATED LINKS

- None
