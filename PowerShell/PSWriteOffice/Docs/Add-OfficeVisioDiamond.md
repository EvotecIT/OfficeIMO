---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Add-OfficeVisioDiamond
## SYNOPSIS
Adds a diamond shape to the current Visio page.

## SYNTAX
### __AllParameterSets
```powershell
Add-OfficeVisioDiamond [[-Text] <string>] [-Page <VisioPage>] [-Key <string>] [-X <double>] [-Y <double>] [-Width <double>] [-Height <double>] [-Unit <VisioMeasurementUnit>] [-Name <string>] [-FillColor <string>] [-LineColor <string>] [-LineWeight <Double>] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Adds a diamond shape to the current Visio page.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> OfficeVisio -Path .\Flow.vsdx {
    VisioDiamond -Key review -Text 'Approved?' -X 4 -Y 4 -Width 1.2 -Height 1 -FillColor '#FEF3C7'
}
```

Adds a decision diamond to the active Visio page.

## PARAMETERS

### -FillColor
Fill color name or hex value.

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

### -Height
Shape height.

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

### -Key
DSL key used by connector commands.

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

### -LineColor
Line color name or hex value.

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

### -LineWeight
Line weight.

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

### -Name
Optional shape name.

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

### -Page
Target page. Optional inside VisioPage or OfficeVisio.

```yaml
Type: VisioPage
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: True (ByValue)
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

### -Text
Text placed inside the shape.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Unit
Measurement unit for coordinates and dimensions.

```yaml
Type: VisioMeasurementUnit
Parameter Sets: __AllParameterSets
Aliases: None
Possible values: Inches, Centimeters, Millimeters

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Width
Shape width.

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

### -X
X coordinate of the shape origin.

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

### -Y
Y coordinate of the shape origin.

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

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `OfficeIMO.Visio.VisioPage`

## OUTPUTS

- `OfficeIMO.Visio.VisioShape`

## RELATED LINKS

- None
