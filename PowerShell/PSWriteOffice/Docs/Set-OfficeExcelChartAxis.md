---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Set-OfficeExcelChartAxis
## SYNOPSIS
Configures common Excel chart axis titles, formats, scale, and gridlines.

## SYNTAX
### Default (Default)
```powershell
Set-OfficeExcelChartAxis [-Chart] <ExcelChart> [-AxisGroup <OfficeChartAxisGroup>] [-CategoryTitle <string>] [-ValueTitle <string>] [-CategoryNumberFormat <string>] [-ValueNumberFormat <string>] [-SourceLinked] [-ValueMinimum <Double>] [-ValueMaximum <Double>] [-ValueMajorUnit <Double>] [-ValueMinorUnit <Double>] [-CategoryMinimum <Double>] [-CategoryMaximum <Double>] [-CategoryMajorUnit <Double>] [-CategoryMinorUnit <Double>] [-ShowCategoryMajorGridlines] [-ShowCategoryMinorGridlines] [-ShowValueMajorGridlines] [-ShowValueMinorGridlines] [-CategoryGridlineColor <OfficeColor>] [-ValueGridlineColor <OfficeColor>] [-GridlineWidth <OfficeLength>] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Configures common Excel chart axis titles, formats, scale, and gridlines.

## EXAMPLES

### EXAMPLE 1
```powershell
New-OfficeExcel -Path .\Reports\Regions.xlsx {
    ExcelSheet 'Summary' {
        ExcelTable (Import-OfficeCsv -Path .\Data\Region-Revenue.csv) -Name RegionSummary
        ExcelChart -Table RegionSummary -Address F2 -Type Column -PassThru | ExcelChartAxis -ValueTitle 'Revenue (USD)'
    }
}
```

Adds a value axis title to a new chart.

### EXAMPLE 2
```powershell
New-OfficeExcel -Path .\Reports\Availability.xlsx {
    ExcelSheet 'Uptime' {
        $chart = ExcelChart -InputObject (Import-OfficeCsv -Path .\Data\Site-Uptime.csv) -Category Site -Value Uptime -Address A1 -Type Column -PassThru
        ExcelChartAxis -Chart $chart -ValueMinimum 0.95 -ValueMaximum 1 -ValueMajorUnit 0.01 -ValueNumberFormat '0%' -CategoryTitle 'Site'
    }
}
```

Scales the value axis from 95 to 100 percent with a percentage format, so small differences in site uptime are visible.

## PARAMETERS

### -AxisGroup
Axis group to configure.

```yaml
Type: OfficeChartAxisGroup
Parameter Sets: Default
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
Type: OfficeColor
Parameter Sets: Default
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
Parameter Sets: Default
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
Parameter Sets: Default
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
Parameter Sets: Default
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
Parameter Sets: Default
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
Parameter Sets: Default
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
Parameter Sets: Default
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
Parameter Sets: Default
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -GridlineWidth
Optional gridline width in points.

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
Emit the object created or changed by the command.

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

### -ShowCategoryMajorGridlines
Show category major gridlines.

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

### -ShowCategoryMinorGridlines
Show category minor gridlines.

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

### -ShowValueMajorGridlines
Show value major gridlines.

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

### -ShowValueMinorGridlines
Show value minor gridlines.

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

### -SourceLinked
Keep axis number formats linked to source cells.

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

### -ValueGridlineColor
Optional value gridline color. Named colors and hexadecimal values are accepted.

```yaml
Type: OfficeColor
Parameter Sets: Default
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
Parameter Sets: Default
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
Parameter Sets: Default
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
Parameter Sets: Default
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
Parameter Sets: Default
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
Parameter Sets: Default
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

- `OfficeIMO.Excel.ExcelChart`

## OUTPUTS

- `OfficeIMO.Excel.ExcelChart`

## RELATED LINKS

- None
