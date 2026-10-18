---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Add-OfficePowerPointImage
## SYNOPSIS
Adds an image to a PowerPoint slide.

## SYNTAX
### __AllParameterSets
```powershell
Add-OfficePowerPointImage [-Path] <string> [-Slide <PowerPointSlide>] [-X <double>] [-Y <double>] [-Width <double>] [-Height <double>] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Places the picture at the requested coordinates using point measurements.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $image = '.\Tests\Assets\CellImage.png'
New-OfficePowerPoint -Path .\Examples\Documents\PowerPointImage.pptx {
    $slide = Add-OfficePowerPointSlide -Layout 1 -PassThru
    Set-OfficePowerPointSlideTitle -Slide $slide -Title 'Evidence'
    Add-OfficePowerPointImage -Slide $slide -Path $image -X 60 -Y 130 -Width 180 -Height 120
}
```

Adds a picture to a generated slide.

## PARAMETERS

### -Height
Image height in points.

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

### -Path
Path to the image file.

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

### -Slide
Target slide that will receive the picture (optional inside DSL).

```yaml
Type: PowerPointSlide
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -Width
Image width in points.

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
Left offset (in points) from the slide origin.

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
Top offset (in points) from the slide origin.

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

- `OfficeIMO.PowerPoint.PowerPointSlide`

## OUTPUTS

- `OfficeIMO.PowerPoint.PowerPointPicture`

## RELATED LINKS

- None
