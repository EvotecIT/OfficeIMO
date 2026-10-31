---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Add-OfficePowerPointLocation
## SYNOPSIS
Adds a named location to a coverage slide.

## SYNTAX
### __AllParameterSets
```powershell
Add-OfficePowerPointLocation [-Name] <string> -MapLeft <double> -MapTop <double> [-Detail <string>] [<CommonParameters>]
```

## DESCRIPTION
Adds a named location to a coverage slide.

## EXAMPLES

### EXAMPLE 1
```powershell
Add-OfficePowerPointLocation -Name 'Name' -MapLeft 1 -MapTop 1
```


## PARAMETERS

### -Detail
Gets or sets the location detail.

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

### -MapLeft
Gets or sets the horizontal fraction of the map.

```yaml
Type: Double
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -MapTop
Gets or sets the vertical fraction of the map.

```yaml
Type: Double
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Name
Gets or sets the location name.

```yaml
Type: String
Parameter Sets: __AllParameterSets
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
