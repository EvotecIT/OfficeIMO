---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Set-OfficeExcelChartAxis
## SYNOPSIS
Configures common Excel chart axis titles, formats, scale, and gridlines.

## SYNTAX
### __AllParameterSets
```powershell
Set-OfficeExcelChartAxis -Chart <ExcelChart> [-AxisGroup <OfficeChartAxisGroup>] [-CategoryTitle <string>] [-ValueTitle <string>] [-CategoryNumberFormat <string>] [-ValueNumberFormat <string>] [-SourceLinked <bool>] [-ValueMinimum <Double>] [-ValueMaximum <Double>] [-ValueMajorUnit <Double>] [-ValueMinorUnit <Double>] [-CategoryMinimum <Double>] [-CategoryMaximum <Double>] [-CategoryMajorUnit <Double>] [-CategoryMinorUnit <Double>] [-ShowCategoryMajorGridlines] [-ShowCategoryMinorGridlines] [-ShowValueMajorGridlines] [-ShowValueMinorGridlines] [-CategoryGridlineColor <string>] [-ValueGridlineColor <string>] [-GridlineWidthPoints <Double>] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Configures common Excel chart axis titles, formats, scale, and gridlines.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $chart | Set-OfficeExcelChartAxis -CategoryTitle 'Month' -ValueTitle 'Revenue' -ValueNumberFormat '$#,##0' -ValueMinimum 0 -ValueMajorUnit 100 -ShowValueMajorGridlines
```

Sets axis titles, value formatting, scale, and major value gridlines.

### EXAMPLE 2
```powershell
PS> $chart |
                Set-OfficeExcelChartAxis -CategoryTitle 'Week' -CategoryNumberFormat 'yyyy-mm-dd' -CategoryMinimum 46000 -CategoryMaximum 46090 -CategoryMajorUnit 14 -CategoryMinorUnit 7
```

Configures the chart's X axis through OfficeIMO's category/date-axis scale support while keeping the workbook chart object on the pipeline.

## PARAMETERS

### -AxisGroup
Axis group to configure.

```yaml
Type: OfficeChartAxisGroup
Parameter Sets: __AllParameterSets
Aliases: None
Possible values: Primary, Secondary

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -CategoryGridlineColor
Optional category gridline color. Named colors and hexadecimal values are accepted.

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

### -CategoryMajorUnit
Category/date axis major unit.

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

### -CategoryMaximum
Category/date axis maximum.

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

### -CategoryMinimum
Category/date axis minimum.

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

### -CategoryMinorUnit
Category/date axis minor unit.

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

### -CategoryNumberFormat
Category axis number format code.

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

### -CategoryTitle
Category axis title.

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

### -Chart
Chart to update.

```yaml
Type: ExcelChart
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -GridlineWidthPoints
Optional gridline width in points.

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

### -PassThru
Emit the object created or changed by the command.

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

### -ShowCategoryMajorGridlines
Show category major gridlines.

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

### -ShowCategoryMinorGridlines
Show category minor gridlines.

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

### -ShowValueMajorGridlines
Show value major gridlines.

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

### -ShowValueMinorGridlines
Show value minor gridlines.

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

### -SourceLinked
Keep axis number formats linked to source cells.

```yaml
Type: Boolean
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ValueGridlineColor
Optional value gridline color. Named colors and hexadecimal values are accepted.

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

### -ValueMajorUnit
Value axis major unit.

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

### -ValueMaximum
Value axis maximum.

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

### -ValueMinimum
Value axis minimum.

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

### -ValueMinorUnit
Value axis minor unit.

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

### -ValueNumberFormat
Value axis number format code.

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

### -ValueTitle
Value axis title.

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

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `OfficeIMO.Excel.ExcelChart`

## OUTPUTS

- `OfficeIMO.Excel.ExcelChart`

## RELATED LINKS

- None
