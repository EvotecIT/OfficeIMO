---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Set-OfficePowerPointBackground
## SYNOPSIS
Sets the slide background color or image.

## SYNTAX
### Color (Default)
```powershell
Set-OfficePowerPointBackground [-Color] <string> [-Slide <PowerPointSlide>] [-PassThru] [<CommonParameters>]
```

### Image
```powershell
Set-OfficePowerPointBackground [-ImagePath] <string> [-Slide <PowerPointSlide>] [-PassThru] [<CommonParameters>]
```

### Clear
```powershell
Set-OfficePowerPointBackground -Clear [-Slide <PowerPointSlide>] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Sets the slide background color or image.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> Get-OfficePowerPointSlide -Presentation $ppt -Index 0 | Set-OfficePowerPointBackground -Color '#F4F7FB'
```

Applies a solid color fill to the slide background.

### EXAMPLE 2
```powershell
PS> Set-OfficePowerPointBackground -Slide $slide -ImagePath '.\hero.png'
```

Uses the provided image as the slide background.

## PARAMETERS

### -Clear
Clears any explicit background color or image.

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
Background color (hex or named color).

```yaml
Type: String
Parameter Sets: Color
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ImagePath
Path to a background image file.

```yaml
Type: String
Parameter Sets: Image
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PassThru
Emit the object created or changed by the command.

```yaml
Type: SwitchParameter
Parameter Sets: Color, Image, Clear
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Slide
Slide to update (optional inside a slide DSL scope).

```yaml
Type: PowerPointSlide
Parameter Sets: Color, Image, Clear
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `OfficeIMO.PowerPoint.PowerPointSlide`

## OUTPUTS

- `OfficeIMO.PowerPoint.PowerPointSlide`

## RELATED LINKS

- None
