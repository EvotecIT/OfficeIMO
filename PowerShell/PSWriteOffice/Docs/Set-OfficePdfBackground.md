---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Set-OfficePdfBackground
## SYNOPSIS
Sets or clears color, image and shape backgrounds in the enclosing PDF composition.

## SYNTAX
### Color (Default)
```powershell
Set-OfficePdfBackground [-Color] <OfficeColor> [<CommonParameters>]
```

### Image
```powershell
Set-OfficePdfBackground -ImagePath <string> [-Fit <OfficeImageFit>] [-Opacity <Double>] [<CommonParameters>]
```

### Clear
```powershell
Set-OfficePdfBackground -Clear [-Target <OfficePdfBackgroundTarget>] [<CommonParameters>]
```

## DESCRIPTION
Sets or clears color, image and shape backgrounds in the enclosing PDF composition.

## EXAMPLES

### EXAMPLE 1
```powershell
Set-OfficePdfBackground -Color 'Value'
```


### EXAMPLE 2
```powershell
Set-OfficePdfBackground -Clear
```


### EXAMPLE 3
```powershell
Set-OfficePdfBackground -ImagePath 'C:\Path'
```


## PARAMETERS

### -Clear
Clear the selected background decoration.

```yaml
Type: SwitchParameter
Parameter Sets: Clear
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Color
Background color.

```yaml
Type: OfficeColor
Parameter Sets: Color
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Fit
How the background image fits the page.

```yaml
Type: OfficeImageFit
Parameter Sets: Image
Aliases: None
Possible values: Stretch, Contain, Cover

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ImagePath
Literal background image path.

```yaml
Type: String
Parameter Sets: Image
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Opacity
Background image opacity.

```yaml
Type: Double
Parameter Sets: Image
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Target
Background decorations to clear.

```yaml
Type: OfficePdfBackgroundTarget
Parameter Sets: Clear
Aliases: None
Possible values: All, Color, Image, Shapes

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
