---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Set-OfficePowerPointSlideSize
## SYNOPSIS
Sets the slide size for a PowerPoint presentation.

## SYNTAX
### Preset (Default)
```powershell
Set-OfficePowerPointSlideSize -Preset <PowerPointSlideSizePreset> [-Presentation <PowerPointPresentation>] [-Portrait] [-PassThru] [<CommonParameters>]
```

### Centimeters
```powershell
Set-OfficePowerPointSlideSize -WidthCm <double> -HeightCm <double> [-Presentation <PowerPointPresentation>] [-PassThru] [<CommonParameters>]
```

### Inches
```powershell
Set-OfficePowerPointSlideSize -WidthInches <double> -HeightInches <double> [-Presentation <PowerPointPresentation>] [-PassThru] [<CommonParameters>]
```

### Points
```powershell
Set-OfficePowerPointSlideSize -WidthPoints <double> -HeightPoints <double> [-Presentation <PowerPointPresentation>] [-PassThru] [<CommonParameters>]
```

### Emus
```powershell
Set-OfficePowerPointSlideSize -WidthEmus <long> -HeightEmus <long> [-Presentation <PowerPointPresentation>] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Supports common presets as well as explicit width and height in centimeters, inches, points, or EMUs.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> New-OfficePowerPoint -Path .\Examples\Documents\PowerPointWidescreen.pptx {
    Set-OfficePowerPointSlideSize -Preset Screen16x9
    Add-OfficePowerPointSlide -Layout 1 -PassThru | Set-OfficePowerPointSlideTitle -Title 'Widescreen deck'
}
```

Applies the 16:9 widescreen preset before adding slides.

### EXAMPLE 2
```powershell
PS> $ppt = New-OfficePowerPoint -Path .\Examples\Documents\PowerPointCustomSize.pptx -NoSave
Set-OfficePowerPointSlideSize -Presentation $ppt -WidthCm 25.4 -HeightCm 14.0
Add-OfficePowerPointSlide -Presentation $ppt -Layout 1 -PassThru | Set-OfficePowerPointSlideTitle -Title 'Custom size'
$ppt | Close-OfficePowerPoint -Save
```

Sets the presentation slide size to a custom 25.4 x 14.0 cm layout.

## PARAMETERS

### -HeightCm
Custom slide height in centimeters.

```yaml
Type: Double
Parameter Sets: Centimeters
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -HeightEmus
Custom slide height in EMUs.

```yaml
Type: Int64
Parameter Sets: Emus
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -HeightInches
Custom slide height in inches.

```yaml
Type: Double
Parameter Sets: Inches
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -HeightPoints
Custom slide height in points.

```yaml
Type: Double
Parameter Sets: Points
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PassThru
Emit the object created or changed by the command.

```yaml
Type: SwitchParameter
Parameter Sets: Preset, Centimeters, Inches, Points, Emus
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Portrait
Apply the preset in portrait orientation.

```yaml
Type: SwitchParameter
Parameter Sets: Preset
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Presentation
Presentation to update (optional inside New-OfficePowerPoint).

```yaml
Type: PowerPointPresentation
Parameter Sets: Preset, Centimeters, Inches, Points, Emus
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -Preset
Preset slide size to apply.

```yaml
Type: PowerPointSlideSizePreset
Parameter Sets: Preset
Aliases: None
Possible values: Screen4x3, Screen16x9, Screen16x10

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -WidthCm
Custom slide width in centimeters.

```yaml
Type: Double
Parameter Sets: Centimeters
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -WidthEmus
Custom slide width in EMUs.

```yaml
Type: Int64
Parameter Sets: Emus
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -WidthInches
Custom slide width in inches.

```yaml
Type: Double
Parameter Sets: Inches
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -WidthPoints
Custom slide width in points.

```yaml
Type: Double
Parameter Sets: Points
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `OfficeIMO.PowerPoint.PowerPointPresentation`

## OUTPUTS

- `OfficeIMO.PowerPoint.PowerPointSlideSize`

## RELATED LINKS

- None
