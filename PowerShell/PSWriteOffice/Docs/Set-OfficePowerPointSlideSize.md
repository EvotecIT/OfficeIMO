---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Set-OfficePowerPointSlideSize
## SYNOPSIS
Sets a presentation's slide size using a preset or physical lengths.

## SYNTAX
### Preset (Default)
```powershell
Set-OfficePowerPointSlideSize -Preset <PowerPointSlideSizePreset> [-Portrait] [-PassThru] [-Document <PowerPointPresentation>] [<CommonParameters>]
```

### Size
```powershell
Set-OfficePowerPointSlideSize -Width <OfficeLength> -Height <OfficeLength> [-PassThru] [-Document <PowerPointPresentation>] [<CommonParameters>]
```

## DESCRIPTION
Numbers mean points; length strings accept pt, cm, mm, in, px or emu. PassThru returns the native slide size.

## EXAMPLES

### EXAMPLE 1
```powershell
Set-OfficePowerPointSlideSize -Document $deck -Preset Screen4x3 -Portrait
```


### EXAMPLE 2
```powershell
New-OfficePowerPoint -Path ./Review.pptx {
    Set-OfficePowerPointSlideSize -Width '25.4cm' -Height '14cm'
    PowerPointSlide -Title 'Service review'
}
```


## PARAMETERS

### -Document
Gets or sets an explicit live document; omission uses the enclosing content block.

```yaml
Type: PowerPointPresentation
Parameter Sets: Preset, Size
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -Height
Gets or sets the custom slide height; numbers mean points.

```yaml
Type: OfficeLength
Parameter Sets: Size
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PassThru
Gets or sets whether to return the native slide size after editing.

```yaml
Type: SwitchParameter
Parameter Sets: Preset, Size
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Portrait
Gets or sets whether to rotate the chosen preset to portrait orientation.

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

### -Preset
Gets or sets a standard slide size.

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

### -Width
Gets or sets the custom slide width; numbers mean points.

```yaml
Type: OfficeLength
Parameter Sets: Size
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
