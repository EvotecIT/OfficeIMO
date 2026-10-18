---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Add-OfficePowerPointBullets
## SYNOPSIS
Adds a bulleted list to a PowerPoint slide.

## SYNTAX
### __AllParameterSets
```powershell
Add-OfficePowerPointBullets [-Bullets] <string[]> [-Slide <PowerPointSlide>] [-X <double>] [-Y <double>] [-Width <double>] [-Height <double>] [-Level <int>] [-BulletChar <string>] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Creates a textbox and populates it with bullet paragraphs.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> New-OfficePowerPoint -Path .\Examples\Documents\PowerPointBullets.pptx {
    $slide = Add-OfficePowerPointSlide -Layout 1 -PassThru
    Set-OfficePowerPointSlideTitle -Slide $slide -Title 'Delivery update'
    Add-OfficePowerPointBullets -Slide $slide -Bullets 'Wins','Risks','Next steps' -X 60 -Y 120 -Width 420 -Height 180
}
```

Creates a slide with a titled bullet list.

## PARAMETERS

### -BulletChar
Optional bullet character (defaults to •).

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

### -Bullets
Bullet items to render.

```yaml
Type: String[]
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Height
Textbox height in points.

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

### -Level
List level (0-8).

```yaml
Type: Int32
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

### -Slide
Target slide that will receive the bullet list (optional inside DSL).

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
Textbox width in points.

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

- `OfficeIMO.PowerPoint.PowerPointTextBox`

## RELATED LINKS

- None
