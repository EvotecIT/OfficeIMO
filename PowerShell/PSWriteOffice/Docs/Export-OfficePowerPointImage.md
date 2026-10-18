---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Export-OfficePowerPointImage
## SYNOPSIS
Exports presentation slides as PNG or SVG images with one result per slide.

## SYNTAX
### Path (Default)
```powershell
Export-OfficePowerPointImage [-Path] <string> [-OutputPath] <string> [-Format <OfficeImageExportFormat>] [-Options <PowerPointPresentationImageExportOptions>] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### Presentation
```powershell
Export-OfficePowerPointImage [-OutputPath] <string> -Presentation <PowerPointPresentation> [-Format <OfficeImageExportFormat>] [-Options <PowerPointPresentationImageExportOptions>] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Exports presentation slides as PNG or SVG images with one result per slide.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> Export-OfficePowerPointImage -Path .\Deck.pptx -OutputPath .\Slides -Format Svg
```

Writes one image per selected slide. Add -PassThru to receive the structured export results.

## PARAMETERS

### -Format
Output image format.

```yaml
Type: OfficeImageExportFormat
Parameter Sets: Path, Presentation
Aliases: None
Possible values: Png, Svg, Jpeg, Tiff, Webp, Bmp, Pbm, Tga, Icon

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Options
Optional slide selection, size, scale, theme, and rendering settings.

```yaml
Type: PowerPointPresentationImageExportOptions
Parameter Sets: Path, Presentation
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -OutputPath
Destination folder.

```yaml
Type: String
Parameter Sets: Path, Presentation
Aliases: None
Possible values:

Required: True
Position: 1
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PassThru
Emit one structured image export result per saved slide.

```yaml
Type: SwitchParameter
Parameter Sets: Path, Presentation
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Path
Path to the presentation.

```yaml
Type: String
Parameter Sets: Path
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Presentation
Open presentation instance.

```yaml
Type: PowerPointPresentation
Parameter Sets: Presentation
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `OfficeIMO.PowerPoint.PowerPointPresentation`

## OUTPUTS

- `OfficeIMO.Drawing.OfficeImageExportResult`

## RELATED LINKS

- None
