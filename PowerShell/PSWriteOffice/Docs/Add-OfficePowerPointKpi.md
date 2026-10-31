---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Add-OfficePowerPointKpi
## SYNOPSIS
Adds a metric to the current designed slide.

## SYNTAX
### Value (Default)
```powershell
Add-OfficePowerPointKpi [-Value] <string> [-Label] <string> [-Change <string>] [-Trend <PowerPointKpiTrend>] [-GoodDirection <PowerPointKpiDirection>] [-Color <OfficeColor>] [<CommonParameters>]
```

### Delta
```powershell
Add-OfficePowerPointKpi [-Value] <string> [-Label] <string> -Delta <double> [-GoodDirection <PowerPointKpiDirection>] [-DeltaUnit <string>] [-Color <OfficeColor>] [<CommonParameters>]
```

## DESCRIPTION
Adds a metric to the current designed slide.

## EXAMPLES

### EXAMPLE 1
```powershell
Add-OfficePowerPointKpi -Value 'Value' -Label 'Value'
```


### EXAMPLE 2
```powershell
Add-OfficePowerPointKpi -Value 'Value' -Label 'Value' -Delta 1
```


## PARAMETERS

### -Change
Gets or sets the change.

```yaml
Type: String
Parameter Sets: Value
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Color
Gets or sets the color.

```yaml
Type: OfficeColor
Parameter Sets: Value, Delta
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Delta
Gets or sets the numeric change; percentages are fractional values.

```yaml
Type: Double
Parameter Sets: Delta
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -DeltaUnit
Gets or sets the change unit, such as %, pt or ms.

```yaml
Type: String
Parameter Sets: Delta
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -GoodDirection
Gets or sets which direction counts as an improvement.

```yaml
Type: PowerPointKpiDirection
Parameter Sets: Value, Delta
Aliases: None
Possible values: Up, Down

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Label
Gets or sets the label.

```yaml
Type: String
Parameter Sets: Value, Delta
Aliases: None
Possible values:

Required: True
Position: 1
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Trend
Gets or sets an explicit change direction.

```yaml
Type: PowerPointKpiTrend
Parameter Sets: Value
Aliases: None
Possible values: Up, Down, Flat

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Value
Gets or sets the value.

```yaml
Type: String
Parameter Sets: Value, Delta
Aliases: None
Possible values:

Required: True
Position: 0
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
