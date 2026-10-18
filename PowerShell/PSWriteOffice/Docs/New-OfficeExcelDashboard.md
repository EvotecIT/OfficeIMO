---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# New-OfficeExcelDashboard
## SYNOPSIS
Builds a worksheet dashboard from tabular data using OfficeIMO dashboard defaults.

## SYNTAX
### Context (Default)
```powershell
New-OfficeExcelDashboard [-InputObject] <Object> [-Title <string>] [-Subtitle <string>] [-TableName <string>] [-TableRow <int>] [-TableColumn <int>] [-TableStyle <ExcelTableStyle>] [-NoAutoFilter] [-NoAutoFit] [-NoChart] [-ChartPreset <ExcelDashboardChartPreset>] [-ChartTitle <string>] [-ChartRow <Int32>] [-ChartColumn <Int32>] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### Path
```powershell
New-OfficeExcelDashboard [-InputObject] <Object> -Path <string> [-Sheet <string>] [-SheetIndex <Int32>] [-Title <string>] [-Subtitle <string>] [-TableName <string>] [-TableRow <int>] [-TableColumn <int>] [-TableStyle <ExcelTableStyle>] [-NoAutoFilter] [-NoAutoFit] [-NoChart] [-ChartPreset <ExcelDashboardChartPreset>] [-ChartTitle <string>] [-ChartRow <Int32>] [-ChartColumn <Int32>] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### Document
```powershell
New-OfficeExcelDashboard [-InputObject] <Object> -Document <ExcelDocument> [-Sheet <string>] [-SheetIndex <Int32>] [-Title <string>] [-Subtitle <string>] [-TableName <string>] [-TableRow <int>] [-TableColumn <int>] [-TableStyle <ExcelTableStyle>] [-NoAutoFilter] [-NoAutoFit] [-NoChart] [-ChartPreset <ExcelDashboardChartPreset>] [-ChartTitle <string>] [-ChartRow <Int32>] [-ChartColumn <Int32>] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Builds a worksheet dashboard from tabular data using OfficeIMO dashboard defaults.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $rows | New-OfficeExcelDashboard -Title 'Sales Dashboard' -TableName Sales -ChartPreset CompactComparison
```

Writes a table and chart into the current Excel DSL worksheet.

## PARAMETERS

### -ChartColumn
Top-left chart column.

```yaml
Type: Int32
Parameter Sets: Context, Path, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ChartPreset
Dashboard chart preset.

```yaml
Type: ExcelDashboardChartPreset
Parameter Sets: Context, Path, Document
Aliases: None
Possible values: Comparison, Trend, Contribution, CompactComparison

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ChartRow
Top-left chart row.

```yaml
Type: Int32
Parameter Sets: Context, Path, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ChartTitle
Chart title. Defaults to Title when omitted.

```yaml
Type: String
Parameter Sets: Context, Path, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Document
Workbook to update outside the DSL context.

```yaml
Type: ExcelDocument
Parameter Sets: Document
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -InputObject
Rows to render in the dashboard table.

```yaml
Type: Object
Parameter Sets: Context, Path, Document
Aliases: Data, DataTable
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -NoAutoFilter
Disable AutoFilter dropdowns on the generated table.

```yaml
Type: SwitchParameter
Parameter Sets: Context, Path, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NoAutoFit
Disable auto-fit for generated table columns.

```yaml
Type: SwitchParameter
Parameter Sets: Context, Path, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NoChart
Do not create a chart.

```yaml
Type: SwitchParameter
Parameter Sets: Context, Path, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PassThru
Emit dashboard build metadata.

```yaml
Type: SwitchParameter
Parameter Sets: Context, Path, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Path
Workbook path to update.

```yaml
Type: String
Parameter Sets: Path
Aliases: InputPath, FilePath
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Sheet
Worksheet name when using Path or Document.

```yaml
Type: String
Parameter Sets: Path, Document
Aliases: SheetName, Worksheet
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -SheetIndex
Worksheet index (0-based) when using Path or Document.

```yaml
Type: Int32
Parameter Sets: Path, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Subtitle
Dashboard subtitle.

```yaml
Type: String
Parameter Sets: Context, Path, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -TableColumn
Top-left column for the generated table.

```yaml
Type: Int32
Parameter Sets: Context, Path, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -TableName
Name for the generated table.

```yaml
Type: String
Parameter Sets: Context, Path, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -TableRow
Top-left row for the generated table.

```yaml
Type: Int32
Parameter Sets: Context, Path, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -TableStyle
Built-in table style.

```yaml
Type: ExcelTableStyle
Parameter Sets: Context, Path, Document
Aliases: None
Possible values: TableStyleLight1, TableStyleLight2, TableStyleLight3, TableStyleLight4, TableStyleLight5, TableStyleLight6, TableStyleLight7, TableStyleLight8, TableStyleLight9, TableStyleLight10, TableStyleLight11, TableStyleLight12, TableStyleLight13, TableStyleLight14, TableStyleLight15, TableStyleLight16, TableStyleLight17, TableStyleLight18, TableStyleLight19, TableStyleLight20, TableStyleLight21, TableStyleMedium1, TableStyleMedium2, TableStyleMedium3, TableStyleMedium4, TableStyleMedium5, TableStyleMedium6, TableStyleMedium7, TableStyleMedium8, TableStyleMedium9, TableStyleMedium10, TableStyleMedium11, TableStyleMedium12, TableStyleMedium13, TableStyleMedium14, TableStyleMedium15, TableStyleMedium16, TableStyleMedium17, TableStyleMedium18, TableStyleMedium19, TableStyleMedium20, TableStyleMedium21, TableStyleMedium22, TableStyleMedium23, TableStyleMedium24, TableStyleMedium25, TableStyleMedium26, TableStyleMedium27, TableStyleMedium28, TableStyleDark1, TableStyleDark2, TableStyleDark3, TableStyleDark4, TableStyleDark5, TableStyleDark6, TableStyleDark7, TableStyleDark8, TableStyleDark9, TableStyleDark10, TableStyleDark11

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Title
Dashboard title.

```yaml
Type: String
Parameter Sets: Context, Path, Document
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

- `System.Object`

## OUTPUTS

- `System.Management.Automation.PSObject`

## RELATED LINKS

- None
